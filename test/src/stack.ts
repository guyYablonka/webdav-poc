import { spawn, type ChildProcess } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { createServer } from 'node:net';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import selfsigned from 'selfsigned';
import { Agent, request } from 'undici';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..');

export interface StackOptions {
  lockMaxTimeoutS: number;
  tokenMaxAgeS?: number;
  resolveCacheMs?: number;
  maxUploadMb?: number;
}

export interface Stack {
  demoUrl: string;
  davUrl: string;
  /** Trusts the stack's self-signed certificate. */
  dispatcher: Agent;
  logs: { demo: string[]; dav: string[] };
  stop(): Promise<void>;
}

async function freePort(): Promise<number> {
  return new Promise((ok, fail) => {
    const server = createServer();
    server.once('error', fail);
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address() as { port: number };
      server.close(() => ok(port));
    });
  });
}

async function makeCert(dir: string): Promise<{ cert: string; key: string; pem: string }> {
  const pems = await selfsigned.generate([{ name: 'commonName', value: 'localhost' }], {
    keySize: 2048,
    algorithm: 'sha256',
    notAfterDate: new Date(Date.now() + 24 * 3600 * 1000),
    extensions: [
      { name: 'basicConstraints', cA: true },
      { name: 'keyUsage', digitalSignature: true, keyEncipherment: true, keyCertSign: true },
      { name: 'extKeyUsage', serverAuth: true },
      {
        name: 'subjectAltName',
        altNames: [
          { type: 2, value: 'localhost' },
          { type: 7, ip: '127.0.0.1' },
        ],
      },
    ],
  });
  const cert = join(dir, 'cert.pem');
  const key = join(dir, 'key.pem');
  writeFileSync(cert, pems.cert);
  writeFileSync(key, pems.private);
  return { cert, key, pem: pems.cert };
}

function launch(name: string, cwd: string, env: Record<string, string>, sink: string[]): ChildProcess {
  const child = spawn(process.execPath, ['--import', 'tsx', 'src/main.ts'], {
    cwd,
    env: { PATH: process.env.PATH ?? '', HOME: process.env.HOME ?? '', ...env },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  const collect = (chunk: Buffer) => sink.push(...chunk.toString('utf8').split('\n').filter(Boolean));
  child.stdout!.on('data', collect);
  child.stderr!.on('data', collect);
  child.on('exit', (code) => sink.push(`[${name} exited with ${code}]`));
  return child;
}

async function waitHealthy(url: string, dispatcher: Agent, logs: string[]): Promise<void> {
  const deadline = Date.now() + 60_000;
  while (Date.now() < deadline) {
    try {
      const res = await request(`${url}/healthz`, { dispatcher });
      await res.body.dump();
      if (res.statusCode === 200) return;
    } catch {
      // not up yet
    }
    if (logs.some((l) => l.includes('exited with'))) break;
    await new Promise((r) => setTimeout(r, 200));
  }
  throw new Error(`${url} did not become healthy:\n${logs.join('\n')}`);
}

async function stopChild(child: ChildProcess): Promise<void> {
  if (child.exitCode !== null) return;
  const exited = new Promise((r) => child.once('exit', r));
  child.kill('SIGTERM');
  const timer = setTimeout(() => child.kill('SIGKILL'), 5000);
  await exited;
  clearTimeout(timer);
}

/** Starts demo-site and dav-service as separate processes over HTTPS, configured only through env. */
export async function startStack(opts: StackOptions): Promise<Stack> {
  const dir = mkdtempSync(join(tmpdir(), 'office-dav-it-'));
  const tls = await makeCert(dir);
  const [demoPort, davPort] = [await freePort(), await freePort()];
  const demoUrl = `https://127.0.0.1:${demoPort}`;
  const davUrl = `https://127.0.0.1:${davPort}`;
  const gatewaySecret = randomBytes(24).toString('hex');
  const common = {
    HOST: '127.0.0.1',
    TLS_CERT: tls.cert,
    TLS_KEY: tls.key,
    GATEWAY_SECRET: gatewaySecret,
    LOCK_MAX_TIMEOUT_S: String(opts.lockMaxTimeoutS),
    MAX_UPLOAD_MB: String(opts.maxUploadMb ?? 64),
    LOG_LEVEL: 'info',
  };
  const logs = { demo: [] as string[], dav: [] as string[] };

  const demo = launch(
    'demo-site',
    join(ROOT, 'apps/demo-site'),
    {
      ...common,
      PORT: String(demoPort),
      DATA_DIR: join(dir, 'data'),
      COOKIE_SECRET: randomBytes(32).toString('hex'),
      DAV_PUBLIC_URL: davUrl,
      TOKEN_MAX_AGE_S: String(opts.tokenMaxAgeS ?? 3600),
      SEED: 'true',
    },
    logs.demo,
  );
  const dav = launch(
    'dav-service',
    join(ROOT, 'services/dav-service'),
    {
      ...common,
      PORT: String(davPort),
      PUBLIC_BASE_URL: davUrl,
      GATEWAY_URL: `${demoUrl}/gw/v1`,
      GATEWAY_CA_FILE: tls.cert,
      RESOLVE_CACHE_MS: String(opts.resolveCacheMs ?? 5000),
      DAV_LOG_HEADERS: 'true',
    },
    logs.dav,
  );

  const dispatcher = new Agent({ connect: { ca: tls.pem } });
  try {
    await Promise.all([waitHealthy(demoUrl, dispatcher, logs.demo), waitHealthy(davUrl, dispatcher, logs.dav)]);
  } catch (err) {
    await Promise.all([stopChild(demo), stopChild(dav)]);
    throw err;
  }

  return {
    demoUrl,
    davUrl,
    dispatcher,
    logs,
    async stop() {
      await Promise.all([stopChild(demo), stopChild(dav)]);
      await dispatcher.close();
    },
  };
}
