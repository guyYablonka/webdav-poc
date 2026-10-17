import type { Readable } from 'node:stream';
import { Agent, request, type Dispatcher } from 'undici';
import {
  FileInfo,
  GatewayErrorBody,
  GatewayHeaders,
  GatewayPaths,
  LockResponse,
  PutContentResponse,
  ResolveResponse,
  type Intent,
  type LockRequest,
} from '@office-dav/gateway-contract';
import { GATEWAY_UNAVAILABLE, GatewayError } from './errors.js';

/** The slice of a zod schema the client uses, so dav-service needs no direct zod dependency. */
interface Schema<T> {
  safeParse(input: unknown): { success: true; data: T } | { success: false; error: unknown };
}

export interface CallContext {
  requestId: string;
  sessionId?: string;
  signal?: AbortSignal;
}

export interface ContentStream {
  body: Readable;
  versionId: string;
  /** Strong, quoted ETag, for example `"<versionId>"`. */
  etag: string;
  contentType: string;
  contentLength: number | undefined;
  lastModified: string | undefined;
}

export interface PutContentOptions {
  /** The token from the If header. Absent when Office saved without one (the gateway decides). */
  lockToken: string | undefined;
  contentLength: number | undefined;
}

/** Everything dav-service needs from the outside world. */
export interface GatewayClient {
  resolve(token: string, intent: Intent, ctx: CallContext): Promise<ResolveResponse>;
  getFile(fileId: string, ctx: CallContext): Promise<FileInfo>;
  getContent(fileId: string, ctx: CallContext): Promise<ContentStream>;
  putContent(fileId: string, body: Readable, opts: PutContentOptions, ctx: CallContext): Promise<PutContentResponse>;
  /** A file this session created with Save As, or null. */
  findSessionFile(sessionId: string, name: string, ctx: CallContext): Promise<FileInfo | null>;
  /** Save As: a PUT to a new name in the session's collection. */
  saveAs(sessionId: string, name: string, body: Readable, opts: { contentLength: number | undefined }, ctx: CallContext): Promise<PutContentResponse>;
  /** Keeps the body of a save whose token no longer resolves. */
  quarantine(token: string, name: string, body: Readable, opts: { contentLength: number | undefined }, ctx: CallContext): Promise<void>;
  lock(req: LockRequest, ctx: CallContext): Promise<LockResponse>;
  unlock(lockToken: string, ctx: CallContext): Promise<void>;
  close(): Promise<void>;
}

/**
 * Supplies the service-to-service auth headers for each gateway call. The POC
 * uses a shared secret; production can return a freshly signed JWT instead.
 */
export type AuthProvider = () => Promise<Record<string, string>>;

export function bearerAuth(secret: string): AuthProvider {
  const headers = { authorization: `Bearer ${secret}` };
  return async () => headers;
}

export interface HttpGatewayClientOptions {
  /** For example `https://localhost:3443/gw/v1`. */
  baseUrl: string;
  authProvider: AuthProvider;
  /** Extra CA certificate(s) to trust for the gateway. */
  ca?: Buffer;
  /** Overrides the default connection pool, mainly for tests. */
  dispatcher?: Dispatcher;
}

const MAX_ERROR_BODY_BYTES = 64 * 1024;

function firstHeader(value: string | string[] | undefined): string | undefined {
  return Array.isArray(value) ? value[0] : value;
}

export class HttpGatewayClient implements GatewayClient {
  private readonly baseUrl: string;
  private readonly dispatcher: Dispatcher;
  private readonly ownsDispatcher: boolean;

  constructor(private readonly opts: HttpGatewayClientOptions) {
    this.baseUrl = opts.baseUrl.replace(/\/+$/, '');
    this.ownsDispatcher = !opts.dispatcher;
    this.dispatcher =
      opts.dispatcher ??
      new Agent({
        connect: opts.ca ? { ca: opts.ca } : {},
        keepAliveTimeout: 10_000,
        headersTimeout: 30_000,
        // Large uploads and downloads stream through; only stalls should time out.
        bodyTimeout: 120_000,
      });
  }

  resolve(token: string, intent: Intent, ctx: CallContext): Promise<ResolveResponse> {
    return this.json('POST', GatewayPaths.resolve, ResolveResponse, 200, ctx, { token, intent });
  }

  getFile(fileId: string, ctx: CallContext): Promise<FileInfo> {
    return this.json('GET', GatewayPaths.file(fileId), FileInfo, 200, ctx);
  }

  async getContent(fileId: string, ctx: CallContext): Promise<ContentStream> {
    const path = GatewayPaths.content(fileId);
    const res = await this.send('GET', path, ctx);
    if (res.statusCode !== 200) throw await this.toError(res, 'GET', path);

    const etag = firstHeader(res.headers.etag);
    const contentType = firstHeader(res.headers['content-type']);
    if (!etag || !contentType) {
      res.body.destroy();
      throw new GatewayError(GATEWAY_UNAVAILABLE, null, `GET ${path}: missing ETag or Content-Type`);
    }
    const length = firstHeader(res.headers['content-length']);
    return {
      body: res.body,
      etag,
      versionId: firstHeader(res.headers[GatewayHeaders.versionId]) ?? etag.replace(/^W\//, '').replace(/"/g, ''),
      contentType,
      contentLength: length === undefined ? undefined : Number(length),
      lastModified: firstHeader(res.headers['last-modified']),
    };
  }

  async putContent(
    fileId: string,
    body: Readable,
    opts: PutContentOptions,
    ctx: CallContext,
  ): Promise<PutContentResponse> {
    const path = GatewayPaths.content(fileId);
    const headers: Record<string, string> = { 'content-type': 'application/octet-stream' };
    if (opts.lockToken) headers[GatewayHeaders.lockToken] = opts.lockToken;
    if (opts.contentLength !== undefined) headers['content-length'] = String(opts.contentLength);
    const res = await this.send('PUT', path, ctx, { body, headers });
    return this.parse(res, 'PUT', path, PutContentResponse, 201);
  }

  async findSessionFile(sessionId: string, name: string, ctx: CallContext): Promise<FileInfo | null> {
    const path = GatewayPaths.sessionFile(sessionId, name);
    const res = await this.send('GET', path, ctx);
    if (res.statusCode === 404) {
      await res.body.dump();
      return null;
    }
    return this.parse(res, 'GET', path, FileInfo, 200);
  }

  async saveAs(
    sessionId: string,
    name: string,
    body: Readable,
    opts: { contentLength: number | undefined },
    ctx: CallContext,
  ): Promise<PutContentResponse> {
    const path = GatewayPaths.sessionFileContent(sessionId, name);
    const headers: Record<string, string> = { 'content-type': 'application/octet-stream' };
    if (opts.contentLength !== undefined) headers['content-length'] = String(opts.contentLength);
    const res = await this.send('PUT', path, ctx, { body, headers });
    return this.parse(res, 'PUT', path, PutContentResponse, 201);
  }

  async quarantine(
    token: string,
    name: string,
    body: Readable,
    opts: { contentLength: number | undefined },
    ctx: CallContext,
  ): Promise<void> {
    const path = GatewayPaths.quarantine;
    const headers: Record<string, string> = {
      'content-type': 'application/octet-stream',
      [GatewayHeaders.davToken]: token,
      [GatewayHeaders.davFileName]: encodeURIComponent(name),
    };
    if (opts.contentLength !== undefined) headers['content-length'] = String(opts.contentLength);
    const res = await this.send('PUT', path, ctx, { body, headers });
    if (res.statusCode !== 201) throw await this.toError(res, 'PUT', path);
    await res.body.dump();
  }

  lock(req: LockRequest, ctx: CallContext): Promise<LockResponse> {
    return this.json('POST', GatewayPaths.locks, LockResponse, 200, ctx, req);
  }

  async unlock(lockToken: string, ctx: CallContext): Promise<void> {
    const path = GatewayPaths.lock(lockToken);
    const res = await this.send('DELETE', path, ctx);
    if (res.statusCode !== 204) throw await this.toError(res, 'DELETE', path);
    await res.body.dump();
  }

  async close(): Promise<void> {
    if (this.ownsDispatcher) await this.dispatcher.close();
  }

  private async json<T>(
    method: 'GET' | 'POST',
    path: string,
    schema: Schema<T>,
    expectStatus: number,
    ctx: CallContext,
    payload?: unknown,
  ): Promise<T> {
    const res = await this.send(method, path, ctx, {
      body: payload === undefined ? undefined : JSON.stringify(payload),
      headers: payload === undefined ? {} : { 'content-type': 'application/json' },
    });
    return this.parse(res, method, path, schema, expectStatus);
  }

  private async parse<T>(
    res: Dispatcher.ResponseData,
    method: string,
    path: string,
    schema: Schema<T>,
    expectStatus: number,
  ): Promise<T> {
    if (res.statusCode !== expectStatus) throw await this.toError(res, method, path);
    let raw: unknown;
    try {
      raw = await res.body.json();
    } catch (err) {
      throw new GatewayError(GATEWAY_UNAVAILABLE, null, `${method} ${path}: response is not JSON`, { cause: err });
    }
    const parsed = schema.safeParse(raw);
    if (!parsed.success) {
      throw new GatewayError(GATEWAY_UNAVAILABLE, null, `${method} ${path}: response violates the contract`, {
        cause: parsed.error,
      });
    }
    return parsed.data;
  }

  private async send(
    method: Dispatcher.HttpMethod,
    path: string,
    ctx: CallContext,
    extra: { body?: string | Readable | undefined; headers?: Record<string, string> } = {},
  ): Promise<Dispatcher.ResponseData> {
    const headers: Record<string, string> = {
      ...(await this.opts.authProvider()),
      [GatewayHeaders.requestId]: ctx.requestId,
      ...extra.headers,
    };
    if (ctx.sessionId) headers[GatewayHeaders.davSession] = ctx.sessionId;
    try {
      return await request(`${this.baseUrl}${path}`, {
        method,
        headers,
        body: extra.body ?? null,
        dispatcher: this.dispatcher,
        signal: ctx.signal ?? null,
      });
    } catch (err) {
      throw new GatewayError(GATEWAY_UNAVAILABLE, null, `${method} ${path}: ${(err as Error).message}`, { cause: err });
    }
  }

  private async toError(res: Dispatcher.ResponseData, method: string, path: string): Promise<GatewayError> {
    let body: GatewayErrorBody | null = null;
    try {
      const text = await readLimited(res.body, MAX_ERROR_BODY_BYTES);
      const parsed = GatewayErrorBody.safeParse(JSON.parse(text));
      if (parsed.success) body = parsed.data;
    } catch {
      // A non-JSON error body still maps by status.
    }
    const detail = body?.reason ?? body?.error ?? 'no body';
    return new GatewayError(res.statusCode, body, `${method} ${path}: gateway answered ${res.statusCode} (${detail})`);
  }
}

async function readLimited(body: Dispatcher.ResponseData['body'], limit: number): Promise<string> {
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of body) {
    size += (chunk as Buffer).length;
    if (size > limit) {
      body.destroy();
      break;
    }
    chunks.push(chunk as Buffer);
  }
  return Buffer.concat(chunks).toString('utf8');
}
