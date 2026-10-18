import type { FastifyPluginAsync, FastifyReply, FastifyRequest } from 'fastify';
import { ZodError } from 'zod';
import {
  GatewayHeaders,
  LockRequest,
  ResolveRequest,
  type FileInfo,
  type GatewayErrorBody,
  type LockResponse,
  type PutContentResponse,
  type ResolveResponse,
} from '@office-dav/gateway-contract';
import type { DavPolicy } from '../domain/dav-policy.js';
import { PolicyError } from '../domain/errors.js';
import type { FileService } from '../domain/files.js';
import type { ServiceAuthenticator } from './service-auth.js';

export interface GatewayDeps {
  policy: DavPolicy;
  files: FileService;
  authenticate: ServiceAuthenticator;
}

function sendError(reply: FastifyReply, status: number, body: GatewayErrorBody): FastifyReply {
  return reply.code(status).send(body);
}

function header(req: FastifyRequest, name: string): string | undefined {
  const value = req.headers[name];
  return typeof value === 'string' && value.trim() ? value.trim() : undefined;
}

function requiredSessionHeader(req: FastifyRequest): string {
  const sessionId = header(req, GatewayHeaders.davSession);
  if (!sessionId) throw new PolicyError(400, 'bad_request', 'X-Dav-Session is required');
  return sessionId;
}

const ABORTED = new Set(['ERR_STREAM_PREMATURE_CLOSE', 'ECONNRESET', 'ERR_STREAM_DESTROYED']);

/**
 * The Gateway API (/gw/v1) that dav-service calls. Self-contained so it can
 * move into the real app as a unit: it needs a DavPolicy, a FileService and
 * a way to authenticate the calling service.
 */
export function gatewayRoutes(deps: GatewayDeps): FastifyPluginAsync {
  const { policy, files } = deps;

  return async (app) => {
    app.addHook('onRequest', async (req, reply) => {
      if (!deps.authenticate(req)) {
        return sendError(reply, 401, { error: 'gateway_auth', message: 'service credentials rejected' });
      }
    });

    // PUT bodies stream straight into the blob store.
    app.addContentTypeParser('application/octet-stream', (_req, _payload, done) => done(null));

    app.setErrorHandler((err, req, reply) => {
      if (err instanceof PolicyError) {
        req.log.info({ status: err.status, code: err.code, detail: err.message }, 'gateway refused');
        return sendError(reply, err.status, { error: err.code, message: err.message, ...err.details });
      }
      if (err instanceof ZodError) {
        return sendError(reply, 400, { error: 'bad_request', message: err.issues.map((i) => `${i.path.join('.')}: ${i.message}`).join('; ') });
      }
      const code = (err as { code?: string }).code;
      if (code && ABORTED.has(code)) {
        req.log.info({ code }, 'upload aborted by the caller');
        return sendError(reply, 400, { error: 'bad_request', message: 'upload aborted' });
      }
      const status = (err as { statusCode?: number }).statusCode;
      if (status && status >= 400 && status < 500) {
        return sendError(reply, status, { error: 'bad_request', message: (err as Error).message });
      }
      req.log.error({ err }, 'gateway failure');
      return sendError(reply, 500, { error: 'internal' });
    });

    app.post('/sessions/resolve', async (req): Promise<ResolveResponse> => {
      const body = ResolveRequest.parse(req.body);
      return policy.resolve(body.token, body.intent);
    });

    app.get<{ Params: { fileId: string } }>('/files/:fileId', async (req): Promise<FileInfo> => {
      const sessionId = header(req, GatewayHeaders.davSession);
      if (sessionId) policy.requireSession(sessionId, req.params.fileId);
      return files.fileInfo(req.params.fileId);
    });

    app.get<{ Params: { fileId: string } }>('/files/:fileId/content', async (req, reply) => {
      const sessionId = header(req, GatewayHeaders.davSession);
      if (sessionId) policy.requireSession(sessionId, req.params.fileId);
      const { file, version, stream } = await files.openVersion(req.params.fileId);
      return reply
        .header('etag', `"${version.id}"`)
        .header(GatewayHeaders.versionId, version.id)
        .header('content-type', file.mime)
        .header('content-length', String(version.size))
        .header('last-modified', new Date(version.created_at).toUTCString())
        .send(stream);
    });

    app.put<{ Params: { fileId: string } }>('/files/:fileId/content', async (req, reply) => {
      const lengthHeader = header(req, 'content-length');
      const result: PutContentResponse = await files.saveFromDav({
        fileId: req.params.fileId,
        sessionId: requiredSessionHeader(req),
        lockToken: header(req, GatewayHeaders.lockToken),
        contentLength: lengthHeader === undefined ? undefined : Number(lengthHeader),
        body: req.raw,
      });
      if (result.revived) req.log.info({ fileId: req.params.fileId }, 'lock revived by PUT');
      if (result.kind !== 'current') req.log.warn({ fileId: req.params.fileId, kind: result.kind }, 'DAV save stored aside');
      return reply.code(201).send(result);
    });

    // Save As support: a file this session created, looked up by the name in Office's URL.
    app.get<{ Params: { sessionId: string; name: string } }>('/sessions/:sessionId/files/:name', async (req): Promise<FileInfo> => {
      policy.requireSession(req.params.sessionId);
      const fileId = policy.findSessionFile(req.params.sessionId, req.params.name);
      if (!fileId) throw new PolicyError(404, 'not_found', 'no such file in this session');
      return files.fileInfo(fileId);
    });

    app.put<{ Params: { sessionId: string; name: string } }>('/sessions/:sessionId/files/:name/content', async (req, reply) => {
      const lengthHeader = header(req, 'content-length');
      const result = await files.saveAsFromDav({
        sessionId: req.params.sessionId,
        name: req.params.name,
        contentLength: lengthHeader === undefined ? undefined : Number(lengthHeader),
        body: req.raw,
      });
      req.log.info({ fileId: result.fileId, kind: result.kind, created: result.created }, 'save as from DAV');
      return reply.code(201).send(result);
    });

    // A save whose token no longer resolves. Keyed by the raw token because there is no live session.
    app.put('/quarantine', async (req, reply) => {
      const token = header(req, GatewayHeaders.davToken);
      if (!token) throw new PolicyError(400, 'bad_request', 'X-Dav-Token is required');
      let fileName = header(req, GatewayHeaders.davFileName) ?? '';
      try {
        fileName = decodeURIComponent(fileName);
      } catch {
        // Keep the raw value; it is only a label for the admin.
      }
      const result = await files.quarantineFromDav({ token, fileName, body: req.raw });
      req.log.warn({ quarantineId: result.quarantineId }, 'DAV save quarantined (session no longer valid)');
      return reply.code(201).send(result);
    });

    app.post('/locks', async (req): Promise<LockResponse> => {
      const result = policy.lock(LockRequest.parse(req.body));
      if (result.revived) req.log.info('lock revived by refresh');
      return result;
    });

    app.delete<{ Params: { lockToken: string } }>('/locks/:lockToken', async (req, reply) => {
      policy.unlock(req.params.lockToken, requiredSessionHeader(req));
      return reply.code(204).send();
    });
  };
}
