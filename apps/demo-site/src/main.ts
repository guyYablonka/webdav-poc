import { buildApp } from './app.js';
import { ConfigError, loadConfig, type DemoConfig } from './config.js';

let config: DemoConfig;
try {
  config = loadConfig();
} catch (err) {
  if (err instanceof ConfigError) {
    console.error(`demo-site: ${err.message}`);
    process.exit(1);
  }
  throw err;
}

const { app } = await buildApp(config);
await app.listen({ host: config.host, port: config.port });
app.log.info(
  {
    dataDir: config.dataDir,
    davPublicUrl: config.davPublicUrl,
    tokenMaxAgeS: config.tokenMaxAgeS,
    lockMaxTimeoutS: config.lockMaxTimeoutS,
  },
  'demo-site ready',
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
