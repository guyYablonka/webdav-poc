import { ConfigError, loadConfig, type DavConfig } from './config.js';
import { buildServer } from './server.js';

let config: DavConfig;
try {
  config = loadConfig();
} catch (err) {
  if (err instanceof ConfigError) {
    console.error(`dav-service: ${err.message}`);
    process.exit(1);
  }
  throw err;
}

const app = buildServer(config);
await app.listen({ host: config.host, port: config.port });
app.log.info(
  {
    publicBaseUrl: config.publicBaseUrl,
    gatewayUrl: config.gatewayUrl,
    tls: Boolean(config.tls),
    lockMaxTimeoutS: config.lockMaxTimeoutS,
    resolveCacheMs: config.resolveCacheMs,
    maxUploadMb: config.maxUploadBytes / 1024 / 1024,
    logHeaders: config.logHeaders,
  },
  'dav-service ready',
);

for (const signal of ['SIGINT', 'SIGTERM'] as const) {
  process.once(signal, () => {
    app.log.info({ signal }, 'shutting down');
    app.close().then(
      () => process.exit(0),
      () => process.exit(1),
    );
  });
}
