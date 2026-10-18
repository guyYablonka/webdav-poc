import { randomUUID } from 'node:crypto';
import { Readable } from 'node:stream';
import type { FileInfo, Intent, LockRequest, LockResponse, PutContentResponse, ResolveResponse } from '@office-dav/gateway-contract';
import type { CallContext, ContentStream, GatewayClient, PutContentOptions } from '../src/gateway/client.js';
import { GatewayError } from '../src/gateway/errors.js';

export const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

async function readAll(body: Readable): Promise<Buffer> {
  const chunks: Buffer[] = [];
  for await (const chunk of body) chunks.push(chunk as Buffer);
  return Buffer.concat(chunks);
}

/** In-memory GatewayClient with just enough lock behaviour to drive the handlers. */
export class FakeGateway implements GatewayClient {
  readonly calls: Array<{ op: string; ctx: CallContext; arg?: unknown }> = [];
  readonly sessions = new Map<string, ResolveResponse>();
  readonly files = new Map<string, { info: FileInfo; content: Buffer }>();
  /** When set, every call fails with this error. */
  failWith: GatewayError | null = null;

  addFile(name: string, content: Buffer | string = 'hello'): FileInfo {
    const fileId = randomUUID();
    const body = Buffer.from(content);
    const info: FileInfo = {
      fileId,
      name,
      mime: DOCX,
      size: body.length,
      versionId: randomUUID(),
      createdAt: '2026-01-01T10:00:00.000Z',
      updatedAt: '2026-01-02T10:00:00.000Z',
      lock: null,
    };
    this.files.set(fileId, { info, content: body });
    return info;
  }

  addSession(fileId: string, opts: Partial<ResolveResponse> = {}): { token: string; session: ResolveResponse } {
    const token = randomUUID().replace(/-/g, '');
    const session: ResolveResponse = {
      sessionId: randomUUID(),
      userId: 'alice',
      displayName: 'Alice',
      fileId,
      canWrite: true,
      state: 'active',
      ...opts,
    };
    this.sessions.set(token, session);
    return { token, session };
  }

  private record(op: string, ctx: CallContext, arg?: unknown): void {
    this.calls.push({ op, ctx, arg });
    if (this.failWith) throw this.failWith;
  }

  private file(fileId: string) {
    const file = this.files.get(fileId);
    if (!file) throw new GatewayError(404, { error: 'not_found' }, 'not found');
    return file;
  }

  async resolve(token: string, intent: Intent, ctx: CallContext): Promise<ResolveResponse> {
    this.record('resolve', ctx, intent);
    const session = this.sessions.get(token);
    if (!session) throw new GatewayError(401, { error: 'session_invalid', reason: 'unknown' }, 'unknown');
    return session;
  }

  async getFile(fileId: string, ctx: CallContext): Promise<FileInfo> {
    this.record('getFile', ctx, fileId);
    return structuredClone(this.file(fileId).info);
  }

  async getContent(fileId: string, ctx: CallContext): Promise<ContentStream> {
    this.record('getContent', ctx, fileId);
    const { info, content } = this.file(fileId);
    return {
      body: Readable.from([content]),
      versionId: info.versionId,
      etag: `"${info.versionId}"`,
      contentType: info.mime,
      contentLength: content.length,
      lastModified: new Date(info.updatedAt).toUTCString(),
    };
  }

  /** Saves stored aside: lock mismatches (conflict) and quarantined bodies. */
  readonly conflicts: Array<{ fileId: string; content: Buffer }> = [];
  readonly quarantined: Array<{ token: string; name: string; content: Buffer }> = [];
  /** Save As files per session id, by lower-cased name. */
  readonly sessionFiles = new Map<string, Map<string, string>>();

  async putContent(fileId: string, body: Readable, opts: PutContentOptions, ctx: CallContext): Promise<PutContentResponse> {
    this.record('putContent', ctx, opts);
    const file = this.file(fileId);
    const content = await readAll(body);
    // dav-service may still hold a cached resolve after a revoke; like the real gateway, the
    // session check happens here too and the body goes to quarantine instead of a 401.
    const revokedToken = ctx.sessionId ? this.revokedSessions.get(ctx.sessionId) : undefined;
    if (revokedToken) {
      this.quarantined.push({ token: revokedToken, name: file.info.name, content });
      return { versionId: randomUUID(), kind: 'quarantined', fileId, size: content.length, updatedAt: new Date().toISOString(), revived: false, created: false };
    }
    // Mirrors the gateway rule: a save is never refused; without the lock it is a conflict copy.
    const lockToken = opts.lockToken ?? (file.info.lock?.sessionId === ctx.sessionId ? file.info.lock?.lockToken : undefined);
    if (!lockToken || file.info.lock?.lockToken !== lockToken) {
      this.conflicts.push({ fileId, content });
      return { versionId: randomUUID(), kind: 'conflict', fileId, size: content.length, updatedAt: new Date().toISOString(), revived: false, created: false };
    }
    file.content = content;
    file.info.size = file.content.length;
    file.info.versionId = randomUUID();
    file.info.updatedAt = new Date().toISOString();
    return { versionId: file.info.versionId, kind: 'current', fileId, size: file.info.size, updatedAt: file.info.updatedAt, revived: false, created: false };
  }

  async findSessionFile(sessionId: string, name: string, ctx: CallContext): Promise<FileInfo | null> {
    this.record('findSessionFile', ctx, name);
    const fileId = this.sessionFiles.get(sessionId)?.get(name.toLowerCase());
    return fileId ? structuredClone(this.file(fileId).info) : null;
  }

  async saveAs(sessionId: string, name: string, body: Readable, _opts: { contentLength: number | undefined }, ctx: CallContext): Promise<PutContentResponse> {
    this.record('saveAs', ctx, name);
    // Unlocked, like the gateway: Office LOCKs the new name itself right after this PUT.
    const info = this.addFile(name, await readAll(body));
    const byName = this.sessionFiles.get(sessionId) ?? new Map<string, string>();
    byName.set(name.toLowerCase(), info.fileId);
    this.sessionFiles.set(sessionId, byName);
    return { versionId: info.versionId, kind: 'current', fileId: info.fileId, size: info.size, updatedAt: info.updatedAt, revived: false, created: true };
  }

  async quarantine(token: string, name: string, body: Readable, _opts: { contentLength: number | undefined }, ctx: CallContext): Promise<void> {
    this.record('quarantine', ctx, name);
    if (!this.revokedTokens.has(token)) {
      body.resume();
      throw new GatewayError(401, { error: 'session_invalid', reason: 'unknown' }, 'unknown');
    }
    this.quarantined.push({ token, name, content: await readAll(body) });
  }

  /** Tokens that resolve to 401 revoked but whose saves may still be quarantined. */
  readonly revokedTokens = new Set<string>();
  /** Revoked session id -> its token, for calls that arrive with a cached session. */
  readonly revokedSessions = new Map<string, string>();

  /** Like revokeSession on the gateway: the token stops resolving and the session's locks go. */
  revoke(token: string): void {
    const session = this.sessions.get(token);
    this.sessions.delete(token);
    this.revokedTokens.add(token);
    if (!session) return;
    this.revokedSessions.set(session.sessionId, token);
    for (const file of this.files.values()) {
      if (file.info.lock?.sessionId === session.sessionId) file.info.lock = null;
    }
  }

  async lock(req: LockRequest, ctx: CallContext): Promise<LockResponse> {
    this.record('lock', ctx, req);
    const file = this.file(req.fileId);
    const expiresAt = new Date(Date.now() + req.timeoutS * 1000).toISOString();
    if (req.lockToken) {
      if (file.info.lock?.lockToken !== req.lockToken) throw new GatewayError(412, { error: 'precondition_failed' }, 'no such lock');
      file.info.lock = { ...file.info.lock, expiresAt, timeoutS: req.timeoutS };
      return { lockToken: req.lockToken, expiresAt, timeoutS: req.timeoutS, revived: false };
    }
    if (file.info.lock) {
      throw new GatewayError(
        423,
        { error: 'locked', heldBy: { userId: 'bob', displayName: 'Bob', since: file.info.updatedAt } },
        'locked',
      );
    }
    const lockToken = `opaquelocktoken:${randomUUID()}`;
    file.info.lock = { lockToken, ownerXml: req.ownerXml ?? null, expiresAt, timeoutS: req.timeoutS, sessionId: req.sessionId };
    return { lockToken, expiresAt, timeoutS: req.timeoutS, revived: false };
  }

  async unlock(lockToken: string, ctx: CallContext): Promise<void> {
    this.record('unlock', ctx, lockToken);
    for (const file of this.files.values()) {
      if (file.info.lock?.lockToken === lockToken) file.info.lock = null;
    }
  }

  async close(): Promise<void> {}
}
