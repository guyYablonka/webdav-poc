import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

export interface DavConfig {
  host: string;
  port: number;
  /** null only when TLS_DISABLED=true (behind a TLS-terminating proxy). */
  tls: { cert: Buffer; key: Buffer } | null;
  /** Base for PROPFIND hrefs, without a trailing slash. */
  publicBaseUrl: string;
  gatewayUrl: string;
  gatewaySecret: string;
  /** Extra CA to trust for the gateway's TLS certificate (for example the mkcert root). */
  gatewayCa: Buffer | undefined;
  maxUploadBytes: number;
  lockMaxTimeoutS: number;
  resolveCacheMs: number;
  logHeaders: boolean;
  logLevel: string;
}

export class ConfigError extends Error {
  override readonly name = 'ConfigError';
}

type Env = Record<string, string | undefined>;

function str(env: Env, name: string, fallback: string): string {
  const value = env[name]?.trim();
  return value ? value : fallback;
}

function int(env: Env, name: string, fallback: number, min: number): number {
  const raw = env[name]?.trim();
  if (!raw) return fallback;
  const value = Number(raw);
  if (!Number.isInteger(value) || value < min) {
    throw new ConfigError(`${name} must be an integer >= ${min}, got "${raw}"`);
  }
  return value;
}

function bool(env: Env, name: string, fallback: boolean): boolean {
  const raw = env[name]?.trim().toLowerCase();
  if (!raw) return fallback;
  return raw === 'true' || raw === '1' || raw === 'yes';
}

function readPem(path: string, name: string): Buffer {
  if (!existsSync(path)) {
    throw new ConfigError(`${name} not found at ${path}. Run "npm run certs" from the repo root, or set ${name}.`);
  }
  return readFileSync(path);
}

export function loadConfig(env: Env = process.env, cwd: string = process.cwd()): DavConfig {
  const gatewaySecret = env.GATEWAY_SECRET?.trim();
  if (!gatewaySecret) throw new ConfigError('GATEWAY_SECRET is required');

  const tlsDisabled = bool(env, 'TLS_DISABLED', false);
  const tls = tlsDisabled
    ? null
    : {
        cert: readPem(resolve(cwd, str(env, 'TLS_CERT', '../../certs/localhost.pem')), 'TLS_CERT'),
        key: readPem(resolve(cwd, str(env, 'TLS_KEY', '../../certs/localhost-key.pem')), 'TLS_KEY'),
      };

  const caFile = env.GATEWAY_CA_FILE?.trim();
  const gatewayCa = caFile ? readPem(resolve(cwd, caFile), 'GATEWAY_CA_FILE') : undefined;

  return {
    host: str(env, 'HOST', 'localhost'),
    port: int(env, 'PORT', 8443, 0),
    tls,
    publicBaseUrl: str(env, 'PUBLIC_BASE_URL', 'https://localhost:8443').replace(/\/+$/, ''),
    gatewayUrl: str(env, 'GATEWAY_URL', 'https://localhost:3443/gw/v1').replace(/\/+$/, ''),
    gatewaySecret,
    gatewayCa,
    maxUploadBytes: int(env, 'MAX_UPLOAD_MB', 512, 1) * 1024 * 1024,
    lockMaxTimeoutS: int(env, 'LOCK_MAX_TIMEOUT_S', 3600, 1),
    resolveCacheMs: int(env, 'RESOLVE_CACHE_MS', 5000, 0),
    logHeaders: bool(env, 'DAV_LOG_HEADERS', false),
    logLevel: str(env, 'LOG_LEVEL', 'info'),
  };
}
