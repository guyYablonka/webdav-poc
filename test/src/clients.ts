import { FormData, request } from 'undici';
import type { Stack } from './stack.js';

export interface DavResponse {
  status: number;
  headers: Record<string, string | string[] | undefined>;
  body: Buffer;
  text: string;
}

export const PROPFIND_OFFICE = `<?xml version="1.0" encoding="utf-8"?>
<D:propfind xmlns:D="DAV:" xmlns:Z="urn:schemas-microsoft-com:"><D:prop>
<D:creationdate/><D:getlastmodified/><D:getcontentlength/><D:getcontenttype/><D:getetag/>
<D:resourcetype/><D:displayname/><D:supportedlock/><D:lockdiscovery/>
<Z:Win32FileAttributes/><Z:Win32LastModifiedTime/></D:prop></D:propfind>`;

export const LOCKINFO = `<?xml version="1.0" encoding="utf-8"?>
<D:lockinfo xmlns:D="DAV:"><D:lockscope><D:exclusive/></D:lockscope><D:locktype><D:write/></D:locktype>
<D:owner><D:href>CONTOSO\\alice</D:href></D:owner></D:lockinfo>`;

export function header(res: DavResponse, name: string): string | undefined {
  const value = res.headers[name];
  return Array.isArray(value) ? value[0] : value;
}

/** Plays the part of Office: arbitrary methods, raw headers and bodies. */
export function davClient(stack: Stack) {
  return async function dav(
    method: string,
    url: string,
    opts: { headers?: Record<string, string>; body?: string | Buffer } = {},
  ): Promise<DavResponse> {
    const res = await request(url, {
      method: method as 'GET',
      headers: { 'user-agent': 'Microsoft Office Word 2014', ...opts.headers },
      body: opts.body ?? null,
      dispatcher: stack.dispatcher,
    });
    const body = Buffer.from(await res.body.arrayBuffer());
    return { status: res.statusCode, headers: res.headers, body, text: body.toString('utf8') };
  };
}

export interface FileSummary {
  id: string;
  name: string;
  versionCount: number;
  lock: { user: { id: string } } | null;
}

export interface FileDetail extends FileSummary {
  versions: Array<{ id: string; versionNo: number; source: string; createdBy: { id: string } }>;
  sessions: Array<{ id: string; user: { id: string }; revokedAt: string | null }>;
}

/** The browser side: login, the JSON API and the admin actions. */
export function demoClient(stack: Stack) {
  async function call<T>(method: string, path: string, cookie: string, body?: { form?: FormData; json?: unknown }): Promise<{ status: number; json: T }> {
    const res = await request(`${stack.demoUrl}${path}`, {
      method: method as 'GET',
      headers: { cookie, 'x-requested-with': 'fetch', ...(body?.json ? { 'content-type': 'application/json' } : {}) },
      body: body?.form ?? (body?.json ? JSON.stringify(body.json) : null),
      dispatcher: stack.dispatcher,
    });
    const text = await res.body.text();
    return { status: res.statusCode, json: (text ? JSON.parse(text) : null) as T };
  }

  return {
    async login(userId: string): Promise<string> {
      const res = await request(`${stack.demoUrl}/login`, {
        method: 'POST',
        headers: { 'content-type': 'application/x-www-form-urlencoded' },
        body: `userId=${userId}`,
        dispatcher: stack.dispatcher,
      });
      await res.body.dump();
      const setCookie = res.headers['set-cookie'];
      const raw = Array.isArray(setCookie) ? setCookie[0]! : setCookie!;
      return raw.split(';')[0]!;
    },

    async logout(cookie: string): Promise<void> {
      const res = await request(`${stack.demoUrl}/logout`, { method: 'POST', headers: { cookie }, dispatcher: stack.dispatcher });
      await res.body.dump();
    },

    async files(cookie: string): Promise<FileSummary[]> {
      return (await call<FileSummary[]>('GET', '/api/files', cookie)).json;
    },

    async file(cookie: string, id: string): Promise<FileDetail> {
      return (await call<FileDetail>('GET', `/api/files/${id}`, cookie)).json;
    },

    async upload(cookie: string, name: string, content: Buffer): Promise<string> {
      const form = new FormData();
      form.append('file', new Blob([content]), name);
      const res = await call<{ id: string }>('POST', '/api/files', cookie, { form });
      if (res.status !== 201) throw new Error(`upload failed: ${res.status}`);
      return res.json.id;
    },

    async uploadVersion(cookie: string, id: string, name: string, content: Buffer): Promise<number> {
      const form = new FormData();
      form.append('file', new Blob([content]), name);
      return (await call('POST', `/api/files/${id}/versions`, cookie, { form })).status;
    },

    async editSession(cookie: string, id: string): Promise<{ davUrl: string; uri: string; sessionId: string; token: string }> {
      const res = await call<{ davUrl: string; uri: string; sessionId: string }>('POST', `/api/files/${id}/edit-session`, cookie);
      if (res.status !== 200) throw new Error(`edit-session failed: ${res.status}`);
      return { ...res.json, token: new URL(res.json.davUrl).pathname.split('/')[2]! };
    },

    async revoke(cookie: string, sessionId: string): Promise<number> {
      return (await call('POST', `/api/sessions/${sessionId}/revoke`, cookie)).status;
    },

    async forceUnlock(cookie: string, id: string): Promise<number> {
      return (await call('POST', `/api/files/${id}/force-unlock`, cookie)).status;
    },
  };
}

export interface SseEvent {
  type: string;
  fileId?: string;
  versionId?: string;
  [key: string]: unknown;
}

/** Subscribes to the demo site's SSE stream and keeps every event. */
export async function openEvents(stack: Stack, cookie: string) {
  const controller = new AbortController();
  const res = await request(`${stack.demoUrl}/events`, {
    headers: { cookie, accept: 'text/event-stream' },
    dispatcher: stack.dispatcher,
    signal: controller.signal,
  });
  if (res.statusCode !== 200) throw new Error(`SSE failed: ${res.statusCode}`);
  const events: SseEvent[] = [];
  const waiters: Array<() => void> = [];
  let buffer = '';
  res.body.setEncoding('utf8');
  res.body.on('data', (chunk: string) => {
    buffer += chunk;
    let cut: number;
    while ((cut = buffer.indexOf('\n\n')) !== -1) {
      const block = buffer.slice(0, cut);
      buffer = buffer.slice(cut + 2);
      const data = block
        .split('\n')
        .filter((l) => l.startsWith('data: '))
        .map((l) => l.slice(6))
        .join('\n');
      if (data) {
        events.push(JSON.parse(data) as SseEvent);
        for (const wake of waiters.splice(0)) wake();
      }
    }
  });
  res.body.on('error', () => {});

  return {
    events,
    async waitFor(predicate: (e: SseEvent) => boolean, timeoutMs = 5000): Promise<SseEvent> {
      const deadline = Date.now() + timeoutMs;
      for (;;) {
        const hit = events.find(predicate);
        if (hit) return hit;
        const left = deadline - Date.now();
        if (left <= 0) throw new Error(`no matching SSE event; got ${JSON.stringify(events.map((e) => e.type))}`);
        await new Promise<void>((wake) => {
          waiters.push(wake);
          setTimeout(wake, left);
        });
      }
    },
    close() {
      controller.abort();
    },
  };
}

export function hrefs(multistatus: string): string[] {
  return [...multistatus.matchAll(/<D:response><D:href>([^<]+)<\/D:href>/g)].map((m) => m[1]!.replace(/&amp;/g, '&'));
}

export const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
