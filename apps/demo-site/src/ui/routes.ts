import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { FastifyPluginAsync, FastifyReply, FastifyRequest } from 'fastify';
import { davFileUrl, officeEditUri } from '@office-dav/gateway-contract';
import type { Auth, AuthUser } from '../auth/auth.js';
import type { DavPolicy } from '../domain/dav-policy.js';
import { PolicyError } from '../domain/errors.js';
import type { FileService } from '../domain/files.js';
import { officeTypeFor } from '../domain/office-types.js';
import { errorPage, fileDetails, filePage, fileRows, filesPage } from './pages.js';
import type { PreviewService } from './preview.js';

export interface UiDeps {
  auth: Auth;
  files: FileService;
  policy: DavPolicy;
  previews: PreviewService;
  davPublicUrl: string;
}

const STATIC_DIR = join(dirname(fileURLToPath(import.meta.url)), 'static');
const STATIC_FILES: Record<string, { type: string; body: Buffer }> = {
  'app.js': { type: 'text/javascript; charset=utf-8', body: readFileSync(join(STATIC_DIR, 'app.js')) },
  'style.css': { type: 'text/css; charset=utf-8', body: readFileSync(join(STATIC_DIR, 'style.css')) },
  'favicon.svg': { type: 'image/svg+xml', body: readFileSync(join(STATIC_DIR, 'favicon.svg')) },
};

declare module 'fastify' {
  interface FastifyRequest {
    user: AuthUser | null;
  }
}

function isApi(req: FastifyRequest): boolean {
  return req.url.startsWith('/api/') || req.url.startsWith('/partials/') || /\/preview(\?|$)/.test(req.url);
}

function contentDisposition(name: string): string {
  const ascii = name.replace(/[^\x20-\x7e]/g, '_').replace(/["\\]/g, '_');
  return `attachment; filename="${ascii}"; filename*=UTF-8''${encodeURIComponent(name)}`;
}

/** Pages, HTML partials for live refresh, and the JSON API the page script calls. */
export function uiRoutes(deps: UiDeps): FastifyPluginAsync {
  const { auth, files, policy, previews } = deps;

  return async (app) => {
    app.decorateRequest('user', null);

    app.addHook('onRequest', async (req, reply) => {
      if (req.url.startsWith('/static/')) return;
      req.user = auth.getCurrentUser(req);
      if (!req.user) {
        if (isApi(req)) return reply.code(401).send({ error: 'unauthorized', message: 'sign in first' });
        return reply.redirect('/login');
      }
      // State-changing API calls must come from our page script (a simple CSRF guard on top of SameSite=Lax).
      if (req.method !== 'GET' && req.headers['x-requested-with'] !== 'fetch') {
        return reply.code(403).send({ error: 'forbidden', message: 'missing X-Requested-With' });
      }
    });

    app.setErrorHandler((err, req, reply) => {
      const status =
        err instanceof PolicyError ? err.status : ((err as { statusCode?: number }).statusCode ?? 500);
      if (status >= 500) req.log.error({ err }, 'request failed');
      const message = status >= 500 ? 'Something went wrong.' : (err as Error).message;
      if (isApi(req) || req.method !== 'GET') return reply.code(status).send({ error: status, message });
      return reply.code(status).type('text/html; charset=utf-8').send(errorPage(req.user, status, message));
    });

    const user = (req: FastifyRequest) => req.user!;
    const htmlReply = (reply: FastifyReply, body: string) => reply.type('text/html; charset=utf-8').send(body);

    app.get<{ Params: { name: string } }>('/static/:name', async (req, reply) => {
      const file = STATIC_FILES[req.params.name];
      if (!file) return reply.code(404).send();
      return reply.type(file.type).header('cache-control', 'no-cache').send(file.body);
    });

    // Pages ------------------------------------------------------------------

    app.get('/', async (req, reply) => htmlReply(reply, filesPage(user(req), files.listSummaries())));

    app.get<{ Params: { id: string } }>('/files/:id', async (req, reply) => {
      const detail = files.detail(req.params.id);
      return htmlReply(reply, filePage(user(req), detail, await previews.render(detail.id)));
    });

    app.get<{ Params: { id: string; versionId: string } }>('/files/:id/versions/:versionId/download', async (req, reply) => {
      const { file, version, stream } = await files.openVersion(req.params.id, req.params.versionId);
      return reply
        .type(file.mime)
        .header('content-length', String(version.size))
        .header('content-disposition', contentDisposition(file.name))
        .send(stream);
    });

    // Partials for live refresh ---------------------------------------------

    app.get('/partials/files', async (_req, reply) => htmlReply(reply, fileRows(files.listSummaries()).value));

    app.get<{ Params: { id: string } }>('/partials/files/:id', async (req, reply) =>
      htmlReply(reply, fileDetails(files.detail(req.params.id)).value),
    );

    app.get<{ Params: { id: string } }>('/files/:id/preview', async (req, reply) =>
      htmlReply(reply, await previews.render(req.params.id)),
    );

    // JSON API ---------------------------------------------------------------

    app.get('/api/me', async (req) => user(req));

    app.get('/api/files', async () => files.listSummaries());

    app.get<{ Params: { id: string } }>('/api/files/:id', async (req) => files.detail(req.params.id));

    app.post('/api/files', async (req, reply) => {
      const part = await req.file();
      if (!part) throw new PolicyError(400, 'bad_request', 'no file in the upload');
      const file = await files.createFile({ name: part.filename, body: part.file, userId: user(req).id, source: 'web' });
      return reply.code(201).send({ id: file.id, name: file.name });
    });

    app.post<{ Params: { id: string } }>('/api/files/:id/versions', async (req, reply) => {
      const part = await req.file();
      if (!part) throw new PolicyError(400, 'bad_request', 'no file in the upload');
      const version = await files.uploadVersion({ fileId: req.params.id, body: part.file, userId: user(req).id });
      return reply.code(201).send({ versionId: version.id, versionNo: version.version_no });
    });

    /** Creates a DAV session and returns the Office URI the browser navigates to. */
    app.post<{ Params: { id: string } }>('/api/files/:id/edit-session', async (req) => {
      const file = files.requireFile(req.params.id);
      const type = officeTypeFor(file.name);
      if (!type) throw new PolicyError(400, 'bad_request', 'not an Office file');
      const { token, session } = policy.createSession({
        userId: user(req).id,
        fileId: file.id,
        userAgent: req.headers['user-agent'] ?? null,
      });
      const davUrl = davFileUrl(deps.davPublicUrl, token, file.id, file.name);
      const lock = policy.activeLock(file.id);
      req.log.info({ fileId: file.id, sessionId: session.id }, 'DAV session created');
      return {
        sessionId: session.id,
        uri: officeEditUri(type.app, davUrl),
        davUrl,
        app: type.label,
        expiresAt: new Date(session.max_expires_at).toISOString(),
        lockedBy: lock && lock.user_id !== user(req).id ? lock.user_id : null,
      };
    });

    app.post<{ Params: { id: string } }>('/api/files/:id/force-unlock', async (req, reply) => {
      files.requireFile(req.params.id);
      const released = policy.forceUnlock(req.params.id);
      req.log.info({ fileId: req.params.id, by: user(req).id, released }, 'force unlock');
      return reply.code(204).send();
    });

    /** Saves kept aside because their session could no longer write (no download here; admin tooling's job). */
    app.get('/api/admin/quarantine', async () =>
      policy.listQuarantine().map((q) => ({
        id: q.id,
        sessionId: q.session_id,
        userId: q.user_id,
        fileId: q.file_id,
        fileName: q.file_name,
        reason: q.reason,
        size: q.size,
        sha256: q.sha256,
        createdAt: new Date(q.created_at).toISOString(),
      })),
    );

    app.post<{ Params: { id: string } }>('/api/sessions/:id/revoke', async (req, reply) => {
      const revoked = policy.revokeSession(req.params.id);
      req.log.info({ sessionId: req.params.id, by: user(req).id, revoked }, 'revoke session');
      return reply.code(204).send();
    });
  };
}
