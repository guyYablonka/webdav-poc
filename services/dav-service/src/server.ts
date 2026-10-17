import { randomUUID } from 'node:crypto';
import Fastify, { LogController, type FastifyInstance, type FastifyReply } from 'fastify';
import type { Logger } from 'pino';
import type { DavConfig } from './config.js';
import { DavError, createDavHandler } from './dav/handler.js';
import { davError } from './dav/xml.js';
import { HttpGatewayClient, bearerAuth, type GatewayClient } from './gateway/client.js';
import { GatewayError, davStatusForGatewayError } from './gateway/errors.js';
import { SessionResolver } from './gateway/session-resolver.js';
import { createLogger, redactHeaders, redactUrl } from './logging.js';

export interface BuildOptions {
  /** Defaults to an HTTP client for GATEWAY_URL authenticated with GATEWAY_SECRET. */
  gateway?: GatewayClient;
  logger?: Logger;
}

/** WebDAV methods Fastify does not route by default, and whether they carry a body. */
const EXTRA_METHODS: ReadonlyArray<readonly [string, boolean]> = [
  ['PROPFIND', true],
  ['PROPPATCH', true],
  ['LOCK', true],
  ['UNLOCK', false],
  ['MKCOL', true],
  ['COPY', false],
  ['MOVE', false],
];
const ROUTED_METHODS = ['OPTIONS', 'GET', 'HEAD', 'PUT', 'DELETE', 'POST', ...EXTRA_METHODS.map(([m]) => m)];

const REQUEST_ID = /^[\w.:-]{1,128}$/;

function sendEmpty(reply: FastifyReply, status: number): FastifyReply {
  return reply.code(status).header('content-length', '0').send();
}

export function buildServer(config: DavConfig, opts: BuildOptions = {}): FastifyInstance {
  const logger = opts.logger ?? createLogger(config.logLevel);
  const gateway =
    opts.gateway ??
    new HttpGatewayClient({
      baseUrl: config.gatewayUrl,
      authProvider: bearerAuth(config.gatewaySecret),
      ...(config.gatewayCa ? { ca: config.gatewayCa } : {}),
    });
  const resolver = new SessionResolver(gateway, config.resolveCacheMs);

  const app = Fastify({
    loggerInstance: logger,
    https: config.tls ? { cert: config.tls.cert, key: config.tls.key } : null,
    // Fastify's own request logs include the raw URL, which carries the token.
    logController: new LogController({ disableRequestLogging: true, requestIdLogLabel: 'reqId' }),
    genReqId: (req) => {
      const incoming = req.headers['x-request-id'];
      return typeof incoming === 'string' && REQUEST_ID.test(incoming) ? incoming : randomUUID();
    },
    exposeHeadRoutes: false,
    return503OnClosing: true,
  }) as unknown as FastifyInstance;

  for (const [method, hasBody] of EXTRA_METHODS) app.addHttpMethod(method, { hasBody });

  // Bodies are read by the handlers: XML with a size cap, PUT streamed to the gateway.
  app.removeAllContentTypeParsers();
  app.addContentTypeParser('*', (_req, _payload, done) => done(null, undefined));

  app.decorateRequest('davSessionId', null);

  app.addHook('onRequest', async (req, reply) => {
    reply.header('x-request-id', req.id);
    if (config.logHeaders) {
      req.log.info({ method: req.method, url: redactUrl(req.url), headers: redactHeaders(req.headers) }, 'dav request');
    }
  });

  app.addHook('onResponse', async (req, reply) => {
    const entry: Record<string, unknown> = {
      method: req.method,
      url: redactUrl(req.url),
      statusCode: reply.statusCode,
      ms: Math.round(reply.elapsedTime),
      userAgent: req.headers['user-agent'],
    };
    if (config.logHeaders) entry.responseHeaders = redactHeaders(reply.getHeaders());
    req.log.info(entry, 'dav response');
  });

  app.setErrorHandler((err, req, reply) => {
    if (err instanceof DavError) {
      req.log[err.status >= 500 ? 'error' : 'info']({ status: err.status, detail: err.message }, 'dav error');
      if (err.closeConnection) reply.header('connection', 'close');
      if (err.xmlBody) return reply.code(err.status).header('content-type', 'application/xml; charset=utf-8').send(err.xmlBody);
      return sendEmpty(reply, err.status);
    }

    if (err instanceof GatewayError) {
      const status = davStatusForGatewayError(err);
      if (status >= 500) {
        req.log.error({ err, gatewayStatus: err.status }, 'gateway unavailable');
      } else if (err.status === 401 && err.body?.error === 'gateway_auth') {
        // Not the user's fault: dav-service's own credentials were rejected.
        req.log.error({ gatewayStatus: 401 }, 'gateway rejected dav-service credentials; check GATEWAY_SECRET');
      } else {
        req.log.info({ status, gatewayStatus: err.status, reason: err.body?.reason ?? err.body?.error }, 'gateway refused');
      }
      if (status === 401) {
        // Office cannot answer a Bearer challenge, so it fails the request instead of prompting.
        reply.header('www-authenticate', 'Bearer realm="office-dav", error="invalid_token"');
      }
      if (status === 423) {
        return reply.code(423).header('content-type', 'application/xml; charset=utf-8').send(davError('no-conflicting-lock'));
      }
      return sendEmpty(reply, status);
    }

    const status = (err as { statusCode?: number }).statusCode ?? 500;
    if (status >= 500) req.log.error({ err }, 'unhandled error');
    else req.log.info({ status, detail: (err as Error).message }, 'request rejected');
    return sendEmpty(reply, status);
  });

  app.setNotFoundHandler((_req, reply) => reply.code(405).header('allow', 'OPTIONS, GET, HEAD, PUT, PROPFIND, PROPPATCH, LOCK, UNLOCK').send());

  app.get('/healthz', async () => ({ status: 'ok' }));

  const handler = createDavHandler({ config, gateway, resolver });
  app.route({ method: ROUTED_METHODS, url: '/', handler });
  app.route({ method: ROUTED_METHODS, url: '/*', handler });

  app.addHook('onClose', async () => {
    await gateway.close();
  });

  return app;
}
