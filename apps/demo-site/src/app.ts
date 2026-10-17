import { mkdirSync } from 'node:fs';
import { join } from 'node:path';
import cookie from '@fastify/cookie';
import formbody from '@fastify/formbody';
import multipart from '@fastify/multipart';
import Fastify, { LogController, type FastifyInstance } from 'fastify';
import { GATEWAY_BASE_PATH } from '@office-dav/gateway-contract';
import { MOCK_USERS, MockOidcAuth } from './auth/mock-oidc-auth.js';
import { authRoutes } from './auth/routes.js';
import type { DemoConfig } from './config.js';
import { DavPolicy, type Clock } from './domain/dav-policy.js';
import { FileService } from './domain/files.js';
import { everyoneCanWrite } from './domain/permissions.js';
import { EventBus } from './events/bus.js';
import { sseRoutes } from './events/sse.js';
import { gatewayRoutes } from './gateway/routes.js';
import { sharedSecretAuth } from './gateway/service-auth.js';
import { seedIfEmpty } from './seed/seed.js';
import { LocalDiskBlobStore } from './storage/blob-store.js';
import { openDatabase } from './storage/db.js';
import { Repo } from './storage/repo.js';
import { PreviewService } from './ui/preview.js';
import { uiRoutes } from './ui/routes.js';

export interface DemoApp {
  app: FastifyInstance;
  repo: Repo;
  policy: DavPolicy;
  files: FileService;
  events: EventBus;
}

export interface BuildOptions {
  clock?: Clock;
  /** Fastify logger setting; defaults to pino at config.logLevel. */
  logger?: boolean | { level: string };
}

export async function buildApp(config: DemoConfig, opts: BuildOptions = {}): Promise<DemoApp> {
  const clock = opts.clock ?? Date.now;
  mkdirSync(config.dataDir, { recursive: true });
  const repo = new Repo(openDatabase(join(config.dataDir, 'demo.sqlite')));
  const blobs = new LocalDiskBlobStore(join(config.dataDir, 'blobs'));
  const events = new EventBus();
  const policy = new DavPolicy(
    repo,
    { tokenMaxAgeS: config.tokenMaxAgeS, lockMaxTimeoutS: config.lockMaxTimeoutS },
    events,
    everyoneCanWrite,
    clock,
  );
  const files = new FileService(repo, blobs, policy, events, { maxUploadBytes: config.maxUploadBytes }, clock);
  const previews = new PreviewService(files, repo);
  const auth = new MockOidcAuth(
    {
      onLogin: (user) => repo.upsertUser(user.id, user.displayName, clock()),
      onLogout: (user) => policy.revokeUserSessions(user.id),
    },
    { secureCookie: config.tls !== null, maxAgeS: 12 * 3600 },
  );

  const app = Fastify({
    logger: opts.logger ?? { level: config.logLevel },
    https: config.tls ? { cert: config.tls.cert, key: config.tls.key } : null,
    requestIdHeader: 'x-request-id',
    // Log gateway traffic and state changes; skip page loads, partials and static files.
    logController: new LogController({
      disableRequestLogging: (req) => req.method === 'GET' && !req.url.startsWith(GATEWAY_BASE_PATH),
    }),
    // SSE connections never go idle; close them on shutdown.
    forceCloseConnections: true,
  }) as unknown as FastifyInstance;

  for (const user of MOCK_USERS) repo.upsertUser(user.id, user.displayName, clock());
  if (config.seed) await seedIfEmpty(files, repo, app.log);

  await app.register(cookie, { secret: config.cookieSecret });
  await app.register(formbody);
  await app.register(multipart, { limits: { fileSize: config.maxUploadBytes, files: 1 } });

  await app.register(gatewayRoutes({ policy, files, authenticate: sharedSecretAuth(config.gatewaySecret) }), {
    prefix: GATEWAY_BASE_PATH,
  });
  await app.register(authRoutes(auth));
  await app.register(sseRoutes({ events, auth }));
  await app.register(uiRoutes({ auth, files, policy, previews, davPublicUrl: config.davPublicUrl }));
  app.get('/healthz', async () => ({ status: 'ok' }));

  app.addHook('onClose', async () => {
    repo.db.close();
  });

  return { app, repo, policy, files, events };
}
