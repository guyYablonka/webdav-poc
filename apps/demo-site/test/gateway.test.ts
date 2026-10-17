import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { FileInfo, GATEWAY_BASE_PATH, LockResponse, PutContentResponse, ResolveResponse } from '@office-dav/gateway-contract';
import { buildApp, type DemoApp } from '../src/app.js';
import type { DemoConfig } from '../src/config.js';

const SECRET = 'test-gateway-secret-0123456789';
const auth = { authorization: `Bearer ${SECRET}` };

let demo: DemoApp;
let fileId: string;

const config = (): DemoConfig => ({
  host: '127.0.0.1',
  port: 0,
  tls: null,
  dataDir: mkdtempSync(join(tmpdir(), 'demo-gw-')),
  gatewaySecret: SECRET,
  cookieSecret: 'cookie-secret-0123456789-0123456789-abcdef',
  davPublicUrl: 'https://dav.test',
  tokenMaxAgeS: 3600,
  lockMaxTimeoutS: 120,
  maxUploadBytes: 1024 * 1024,
  seed: true,
  logLevel: 'silent',
});

beforeAll(async () => {
  demo = await buildApp(config(), { logger: false });
  fileId = demo.files.listSummaries().find((f) => f.name === 'Project Plan.docx')!.id;
});

afterAll(async () => {
  await demo.app.close();
});

const gw = (path: string) => `${GATEWAY_BASE_PATH}${path}`;

async function login(userId: string): Promise<string> {
  const res = await demo.app.inject({ method: 'POST', url: '/login', payload: { userId } });
  expect(res.statusCode).toBe(303);
  const cookie = res.cookies.find((c) => c.name === 'demo_session')!;
  return `${cookie.name}=${cookie.value}`;
}

async function editSession(cookie: string): Promise<{ token: string; davUrl: string; uri: string }> {
  const res = await demo.app.inject({
    method: 'POST',
    url: `/api/files/${fileId}/edit-session`,
    headers: { cookie, 'x-requested-with': 'fetch' },
  });
  expect(res.statusCode).toBe(200);
  const body = res.json() as { davUrl: string; uri: string };
  return { ...body, token: new URL(body.davUrl).pathname.split('/')[2]! };
}

async function resolve(token: string) {
  const res = await demo.app.inject({ method: 'POST', url: gw('/sessions/resolve'), headers: auth, payload: { token, intent: 'read' } });
  return { status: res.statusCode, body: res.json() };
}

describe('gateway authentication', () => {
  it('rejects calls without the service secret', async () => {
    for (const headers of [{}, { authorization: 'Bearer wrong' }, { authorization: SECRET }]) {
      const res = await demo.app.inject({ method: 'GET', url: gw(`/files/${fileId}`), headers });
      expect(res.statusCode).toBe(401);
      expect(res.json()).toEqual({ error: 'gateway_auth', message: 'service credentials rejected' });
    }
  });
});

describe('gateway endpoints', () => {
  it('runs resolve, file info, lock, put, content and unlock against the contract', async () => {
    const cookie = await login('alice');
    const { token, uri } = await editSession(cookie);
    expect(uri).toMatch(/^ms-word:ofe\|u\|https:\/\/dav\.test\/s\/[A-Za-z0-9_-]{43}\/[0-9a-f-]{36}\/Project%20Plan\.docx$/);

    const resolved = await resolve(token);
    expect(resolved.status).toBe(200);
    const session = ResolveResponse.parse(resolved.body);
    expect(session).toMatchObject({ userId: 'alice', displayName: 'Alice', fileId, canWrite: true, state: 'active' });
    const sessionHeaders = { ...auth, 'x-dav-session': session.sessionId };

    const info = FileInfo.parse((await demo.app.inject({ method: 'GET', url: gw(`/files/${fileId}`), headers: sessionHeaders })).json());
    expect(info.lock).toBeNull();

    const lockRes = await demo.app.inject({
      method: 'POST',
      url: gw('/locks'),
      headers: sessionHeaders,
      payload: { fileId, sessionId: session.sessionId, timeoutS: 3600, ownerXml: '<D:owner xmlns:D="DAV:">me</D:owner>' },
    });
    expect(lockRes.statusCode).toBe(200);
    const lock = LockResponse.parse(lockRes.json());
    expect(lock.timeoutS).toBe(120);

    const locked = FileInfo.parse((await demo.app.inject({ method: 'GET', url: gw(`/files/${fileId}`), headers: sessionHeaders })).json());
    expect(locked.lock).toMatchObject({ lockToken: lock.lockToken, sessionId: session.sessionId, ownerXml: '<D:owner xmlns:D="DAV:">me</D:owner>' });

    const putRes = await demo.app.inject({
      method: 'PUT',
      url: gw(`/files/${fileId}/content`),
      headers: { ...sessionHeaders, 'x-lock-token': lock.lockToken, 'content-type': 'application/octet-stream' },
      payload: Buffer.from('new bytes'),
    });
    expect(putRes.statusCode).toBe(201);
    const saved = PutContentResponse.parse(putRes.json());
    expect(saved).toMatchObject({ size: 9, revived: false });

    const content = await demo.app.inject({ method: 'GET', url: gw(`/files/${fileId}/content`), headers: sessionHeaders });
    expect(content.statusCode).toBe(200);
    expect(content.body).toBe('new bytes');
    expect(content.headers.etag).toBe(`"${saved.versionId}"`);
    expect(content.headers['content-type']).toBe('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    expect(content.headers['content-length']).toBe('9');

    const unlock = await demo.app.inject({
      method: 'DELETE',
      url: gw(`/locks/${encodeURIComponent(lock.lockToken)}`),
      headers: sessionHeaders,
    });
    expect(unlock.statusCode).toBe(204);
    expect(demo.repo.getLock(fileId)).toBeUndefined();
  });

  it('answers 401 with a reason for unknown and revoked tokens', async () => {
    expect(await resolve('not-a-token')).toEqual({ status: 401, body: { error: 'session_invalid', message: 'DAV session is unknown', reason: 'unknown' } });
    const cookie = await login('bob');
    const { token } = await editSession(cookie);
    const session = ResolveResponse.parse((await resolve(token)).body);
    // Logging out revokes the user's DAV sessions.
    await demo.app.inject({ method: 'POST', url: '/logout', headers: { cookie } });
    expect((await resolve(token)).body.reason).toBe('revoked');
    const res = await demo.app.inject({
      method: 'GET',
      url: gw(`/files/${fileId}`),
      headers: { ...auth, 'x-dav-session': session.sessionId },
    });
    expect(res.statusCode).toBe(401);
  });

  it('refuses a PUT without a lock before reading the body', async () => {
    const cookie = await login('alice');
    const { token } = await editSession(cookie);
    const session = ResolveResponse.parse((await resolve(token)).body);
    const res = await demo.app.inject({
      method: 'PUT',
      url: gw(`/files/${fileId}/content`),
      headers: { ...auth, 'x-dav-session': session.sessionId, 'content-type': 'application/octet-stream' },
      payload: Buffer.from('x'),
    });
    expect(res.statusCode).toBe(423);
  });

  it('validates request bodies', async () => {
    const res = await demo.app.inject({ method: 'POST', url: gw('/locks'), headers: auth, payload: { fileId } });
    expect(res.statusCode).toBe(400);
    expect(res.json().error).toBe('bad_request');
  });

  it('refuses web uploads while locked with 409', async () => {
    const cookie = await login('alice');
    const { token } = await editSession(cookie);
    const session = ResolveResponse.parse((await resolve(token)).body);
    await demo.app.inject({ method: 'POST', url: gw('/locks'), headers: auth, payload: { fileId, sessionId: session.sessionId, timeoutS: 60 } });
    const boundary = 'xyz';
    const payload = `--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="Project Plan.docx"\r\nContent-Type: application/octet-stream\r\n\r\nweb\r\n--${boundary}--\r\n`;
    const res = await demo.app.inject({
      method: 'POST',
      url: `/api/files/${fileId}/versions`,
      headers: { cookie, 'x-requested-with': 'fetch', 'content-type': `multipart/form-data; boundary=${boundary}` },
      payload,
    });
    expect(res.statusCode).toBe(409);
    expect(res.json().message).toMatch(/being edited by Alice/);
  });
});
