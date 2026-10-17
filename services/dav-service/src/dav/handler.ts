import { Transform, finished, type TransformCallback } from 'node:stream';
import type { FastifyReply, FastifyRequest } from 'fastify';
import {
  davCollectionUrl,
  davFileUrl,
  encodePathSegment,
  type FileInfo,
  type Intent,
  type ResolveResponse,
} from '@office-dav/gateway-contract';
import type { DavConfig } from '../config.js';
import type { CallContext, GatewayClient } from '../gateway/client.js';
import { GatewayError } from '../gateway/errors.js';
import type { SessionResolver } from '../gateway/session-resolver.js';
import { etagMatches, parseDepth, parseIfHeader, parseLockTokenHeader, parseTimeout } from './headers.js';
import { parseTarget, sameFileName } from './paths.js';
import { lockResponseBody, propfindBody, proppatchBody, type DavResource } from './props.js';
import { XmlError, davError, parseLockinfo, parsePropfind, parseProppatch } from './xml.js';

/** A protocol-level failure with the status (and optional XML body) to send. */
export class DavError extends Error {
  override readonly name = 'DavError';

  constructor(
    readonly status: number,
    message: string,
    readonly xmlBody?: string,
    /**
     * Close the connection after replying, so a client stops sending a body we
     * refused. Only for oversized bodies: Node drains small unread bodies itself.
     */
    readonly closeConnection = false,
  ) {
    super(message);
  }
}

export const ALLOW = 'OPTIONS, GET, HEAD, PUT, PROPFIND, PROPPATCH, LOCK, UNLOCK';
const FORBIDDEN_METHODS = new Set(['MOVE', 'COPY', 'DELETE', 'MKCOL']);
const XML_BODY_LIMIT = 1024 * 1024;
const XML_CONTENT_TYPE = 'application/xml; charset=utf-8';

export interface DavHandlerDeps {
  config: DavConfig;
  gateway: GatewayClient;
  resolver: SessionResolver;
  /** When the service started; used as the date of the synthetic collections. */
  startedAt?: Date;
}

declare module 'fastify' {
  interface FastifyRequest {
    /** Set once the token resolves, so logs can name the session instead of the token. */
    davSessionId: string | null;
  }
}

function intentFor(method: string): Intent {
  if (method === 'PUT') return 'write';
  if (method === 'LOCK' || method === 'UNLOCK') return 'lock';
  return 'read';
}

function httpDate(iso: string): string {
  return new Date(iso).toUTCString();
}

function notModifiedSince(header: string | undefined, updatedAt: string): boolean {
  if (!header) return false;
  const since = Date.parse(header);
  if (Number.isNaN(since)) return false;
  return Math.floor(Date.parse(updatedAt) / 1000) <= Math.floor(since / 1000);
}

async function readXmlBody(req: FastifyRequest): Promise<string> {
  const declared = Number(req.headers['content-length'] ?? 0);
  if (declared > XML_BODY_LIMIT) throw new DavError(413, 'XML body too large', undefined, true);
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req.raw) {
    size += (chunk as Buffer).length;
    if (size > XML_BODY_LIMIT) throw new DavError(413, 'XML body too large', undefined, true);
    chunks.push(chunk as Buffer);
  }
  return Buffer.concat(chunks).toString('utf8').replace(/^﻿/, '');
}

function parseXml<T>(parse: (body: string) => T, body: string): T {
  try {
    return parse(body);
  } catch (err) {
    if (err instanceof XmlError) throw new DavError(400, err.message);
    throw err;
  }
}

function sendXml(reply: FastifyReply, status: number, xml: string): FastifyReply {
  return reply.code(status).header('content-type', XML_CONTENT_TYPE).send(xml);
}

function sendEmpty(reply: FastifyReply, status: number): FastifyReply {
  return reply.code(status).header('content-length', '0').send();
}

function sendOptions(reply: FastifyReply): FastifyReply {
  return reply
    .code(200)
    .header('dav', '1,2')
    .header('ms-author-via', 'DAV')
    .header('allow', ALLOW)
    .header('content-length', '0')
    .send();
}

/** Counts PUT bytes on their way to the gateway and fails the stream past the limit. */
class UploadLimit extends Transform {
  bytes = 0;
  exceeded = false;

  constructor(private readonly limit: number) {
    super();
  }

  override _transform(chunk: Buffer, _encoding: BufferEncoding, callback: TransformCallback): void {
    this.bytes += chunk.length;
    if (this.bytes > this.limit) {
      this.exceeded = true;
      callback(new DavError(413, 'upload exceeds MAX_UPLOAD_MB'));
      return;
    }
    callback(null, chunk);
  }
}

export function createDavHandler(deps: DavHandlerDeps) {
  const { config, gateway, resolver } = deps;
  const startedAt = deps.startedAt ?? new Date();
  const base = config.publicBaseUrl;

  function syntheticCollection(href: string, displayName: string): DavResource {
    return { href, isCollection: true, displayName, createdAt: startedAt, modifiedAt: startedAt };
  }

  function fileResource(token: string, file: FileInfo): DavResource {
    return {
      href: davFileUrl(base, token, file.fileId, file.name),
      isCollection: false,
      displayName: file.name,
      createdAt: new Date(file.createdAt),
      modifiedAt: new Date(file.updatedAt),
      file: { size: file.size, mime: file.mime, etag: `"${file.versionId}"`, lock: file.lock },
    };
  }

  /** `/`, `/s/` and `/s/:token/`: just enough for Office's parent-path probes. */
  async function handleSyntheticCollection(req: FastifyRequest, reply: FastifyReply, resource: DavResource) {
    switch (req.method) {
      case 'OPTIONS':
        return sendOptions(reply);
      case 'PROPFIND': {
        const request = parseXml(parsePropfind, await readXmlBody(req));
        // Depth 1 is answered like Depth 0: these collections have nothing to list.
        return sendXml(reply, 207, propfindBody([resource], request));
      }
      case 'PROPPATCH': {
        const names = parseXml(parseProppatch, await readXmlBody(req));
        return sendXml(reply, 207, proppatchBody(resource.href, names));
      }
      default:
        return reply.code(405).header('allow', 'OPTIONS, PROPFIND, PROPPATCH').send();
    }
  }

  async function handleCollection(
    req: FastifyRequest,
    reply: FastifyReply,
    token: string,
    session: ResolveResponse,
    ctx: CallContext,
  ) {
    if (req.method === 'OPTIONS') return sendOptions(reply);
    if (req.method !== 'PROPFIND' && req.method !== 'PROPPATCH') {
      return reply.code(405).header('allow', 'OPTIONS, PROPFIND, PROPPATCH').send();
    }
    const file = await gateway.getFile(session.fileId, ctx);
    const collection: DavResource = {
      href: davCollectionUrl(base, token, file.fileId),
      isCollection: true,
      displayName: file.fileId,
      createdAt: new Date(file.createdAt),
      modifiedAt: new Date(file.updatedAt),
    };
    if (req.method === 'PROPPATCH') {
      const names = parseXml(parseProppatch, await readXmlBody(req));
      return sendXml(reply, 207, proppatchBody(collection.href, names));
    }
    const depth = parseDepth(req.headers.depth as string | undefined, 'infinity');
    if (depth === null) throw new DavError(400, 'invalid Depth header');
    const request = parseXml(parsePropfind, await readXmlBody(req));
    const resources = depth === '0' ? [collection] : [collection, fileResource(token, file)];
    return sendXml(reply, 207, propfindBody(resources, request));
  }

  async function handleGet(req: FastifyRequest, reply: FastifyReply, file: FileInfo, ctx: CallContext) {
    const etag = `"${file.versionId}"`;
    const ifNoneMatch = req.headers['if-none-match'];
    const notModified = ifNoneMatch
      ? etagMatches(ifNoneMatch, etag)
      : notModifiedSince(req.headers['if-modified-since'], file.updatedAt);
    if (notModified) {
      return reply.code(304).header('etag', etag).header('last-modified', httpDate(file.updatedAt)).send();
    }
    if (req.method === 'HEAD') {
      return reply
        .code(200)
        .header('etag', etag)
        .header('last-modified', httpDate(file.updatedAt))
        .header('content-type', file.mime)
        .header('content-length', String(file.size))
        .header('accept-ranges', 'none')
        .send();
    }
    const content = await gateway.getContent(file.fileId, ctx);
    reply
      .code(200)
      .header('etag', content.etag)
      .header('last-modified', content.lastModified ?? httpDate(file.updatedAt))
      .header('content-type', content.contentType)
      .header('accept-ranges', 'none');
    if (content.contentLength !== undefined) reply.header('content-length', String(content.contentLength));
    return reply.send(content.body);
  }

  async function handleLock(
    req: FastifyRequest,
    reply: FastifyReply,
    session: ResolveResponse,
    file: FileInfo,
    href: string,
    ctx: CallContext,
  ) {
    const body = await readXmlBody(req);
    const timeoutS = parseTimeout(req.headers.timeout as string | undefined, config.lockMaxTimeoutS);

    if (body.trim() === '') {
      // Refresh: no body, the lock token comes from the If header.
      const lockToken = parseIfHeader(req.headers.if as string | undefined)[0];
      if (!lockToken) throw new DavError(400, 'lock refresh needs an If header with the lock token');
      const granted = await gateway.lock({ fileId: file.fileId, sessionId: session.sessionId, timeoutS, lockToken }, ctx);
      if (granted.revived) req.log.info({ fileId: file.fileId }, 'lock revived by refresh');
      const current = await gateway.getFile(file.fileId, ctx);
      const ownerXml = current.lock?.lockToken === granted.lockToken ? current.lock.ownerXml : null;
      return sendXml(
        reply,
        200,
        lockResponseBody({ lockToken: granted.lockToken, ownerXml, timeoutS: granted.timeoutS }, href),
      );
    }

    const depth = parseDepth(req.headers.depth as string | undefined, 'infinity');
    if (depth === null || depth === '1') throw new DavError(400, 'LOCK supports Depth 0 or infinity');
    const info = parseXml(parseLockinfo, body);
    if (info.scope === 'shared') throw new DavError(412, 'shared locks are not supported');
    if (!info.isWrite) throw new DavError(400, 'only write locks are supported');

    const granted = await gateway.lock(
      {
        fileId: file.fileId,
        sessionId: session.sessionId,
        timeoutS,
        ...(info.ownerXml ? { ownerXml: info.ownerXml } : {}),
      },
      ctx,
    );
    reply.header('lock-token', `<${granted.lockToken}>`);
    return sendXml(
      reply,
      200,
      lockResponseBody({ lockToken: granted.lockToken, ownerXml: info.ownerXml, timeoutS: granted.timeoutS }, href),
    );
  }

  async function handleUnlock(req: FastifyRequest, reply: FastifyReply, ctx: CallContext) {
    const lockToken = parseLockTokenHeader(req.headers['lock-token'] as string | undefined);
    if (!lockToken) throw new DavError(400, 'UNLOCK needs a Lock-Token header');
    await gateway.unlock(lockToken, ctx);
    return sendEmpty(reply, 204);
  }

  async function handlePut(req: FastifyRequest, reply: FastifyReply, file: FileInfo, ctx: CallContext) {
    const lockToken = parseIfHeader(req.headers.if as string | undefined)[0];
    if (!lockToken) throw new DavError(423, 'PUT requires a lock token', davError('lock-token-submitted'));

    const ifMatch = req.headers['if-match'];
    if (ifMatch && !etagMatches(ifMatch, `"${file.versionId}"`)) {
      throw new DavError(412, 'If-Match does not match the current version');
    }

    const lengthHeader = req.headers['content-length'];
    const contentLength = lengthHeader === undefined ? undefined : Number(lengthHeader);
    if (contentLength !== undefined && contentLength > config.maxUploadBytes) {
      throw new DavError(413, 'upload exceeds MAX_UPLOAD_MB', undefined, true);
    }

    // Stream without buffering. `pipe` (not `pipeline`) keeps the client socket
    // alive if the gateway answers early and drops the body, so we can still
    // send the gateway's status back to Office.
    const upload = new UploadLimit(config.maxUploadBytes);
    req.raw.pipe(upload);
    finished(req.raw, (err) => {
      if (err && !upload.destroyed) upload.destroy(new Error('client aborted the upload'));
    });

    try {
      const saved = await gateway.putContent(file.fileId, upload, { lockToken, contentLength }, ctx);
      if (saved.revived) req.log.info({ fileId: file.fileId }, 'lock revived by PUT');
      return reply.code(204).header('etag', `"${saved.versionId}"`).header('content-length', '0').send();
    } catch (err) {
      if (upload.exceeded) throw new DavError(413, 'upload exceeds MAX_UPLOAD_MB', undefined, true);
      throw err;
    } finally {
      if (!req.raw.readableEnded) {
        // The gateway stopped reading early. Discard the rest of the body.
        req.raw.unpipe(upload);
        req.raw.resume();
      }
    }
  }

  async function handleFile(
    req: FastifyRequest,
    reply: FastifyReply,
    target: { token: string; fileName: string },
    session: ResolveResponse,
    ctx: CallContext,
  ) {
    const file = await gateway.getFile(session.fileId, ctx);
    if (!sameFileName(file.name, target.fileName)) {
      // Only the session's file exists here; creating files is out of scope.
      if (req.method === 'PUT' || req.method === 'LOCK') throw new DavError(403, 'cannot create new resources');
      throw new DavError(404, 'no such file in this collection');
    }
    const resource = fileResource(target.token, file);

    switch (req.method) {
      case 'OPTIONS':
        return sendOptions(reply);
      case 'PROPFIND': {
        if (parseDepth(req.headers.depth as string | undefined, 'infinity') === null) {
          throw new DavError(400, 'invalid Depth header');
        }
        const request = parseXml(parsePropfind, await readXmlBody(req));
        return sendXml(reply, 207, propfindBody([resource], request));
      }
      case 'PROPPATCH': {
        const names = parseXml(parseProppatch, await readXmlBody(req));
        return sendXml(reply, 207, proppatchBody(resource.href, names));
      }
      case 'GET':
      case 'HEAD':
        return handleGet(req, reply, file, ctx);
      case 'LOCK':
        if (!session.canWrite) throw new DavError(403, 'session is read-only');
        return handleLock(req, reply, session, file, resource.href, ctx);
      case 'UNLOCK':
        return handleUnlock(req, reply, ctx);
      case 'PUT':
        if (!session.canWrite) throw new DavError(403, 'session is read-only');
        return handlePut(req, reply, file, ctx);
      default:
        return reply.code(405).header('allow', ALLOW).send();
    }
  }

  return async function handle(req: FastifyRequest, reply: FastifyReply) {
    const target = parseTarget(req.url);
    switch (target.kind) {
      case 'office-probe':
      case 'not-found':
        return sendEmpty(reply, 404);
      case 'bad-request':
        return sendEmpty(reply, 400);
      default:
    }
    if (FORBIDDEN_METHODS.has(req.method)) return sendEmpty(reply, 403);
    if (target.kind === 'root') {
      return handleSyntheticCollection(req, reply, syntheticCollection(`${base}${target.path}`, ''));
    }

    const ctx: CallContext = { requestId: req.id };
    try {
      const session = await resolver.resolve(target.token, intentFor(req.method), ctx);
      ctx.sessionId = session.sessionId;
      req.davSessionId = session.sessionId;
      req.log = req.log.child({ sessionId: session.sessionId });

      if (target.kind === 'session') {
        const href = `${base}/s/${encodePathSegment(target.token)}/`;
        return await handleSyntheticCollection(req, reply, syntheticCollection(href, 'session'));
      }
      if (target.fileId !== session.fileId) throw new DavError(404, 'file is not part of this session');
      if (target.kind === 'collection') return await handleCollection(req, reply, target.token, session, ctx);
      return await handleFile(req, reply, target, session, ctx);
    } catch (err) {
      // Any 401 means the session is gone; drop it so the next request asks again.
      if (err instanceof GatewayError && err.status === 401) resolver.invalidate(target.token);
      throw err;
    }
  };
}
