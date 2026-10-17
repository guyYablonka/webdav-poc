import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

export interface DemoConfig {
  host: string;
  port: number;
  /** null only when TLS_DISABLED=true (tests, or behind a TLS-terminating proxy). */
  tls: { cert: Buffer; key: Buffer } | null;
  /** Holds the SQLite database and the blob store. */
  dataDir: string;
  gatewaySecret: string;
  cookieSecret: string;
  /** dav-service's public base URL, used to build the Office URIs. */
  davPublicUrl: string;
  /** Backstop lifetime of a DAV session token. */
  tokenMaxAgeS: number;
  /** Cap for a granted lock timeout. */
  lockMaxTimeoutS: number;
  maxUploadBytes: number;
  seed: boolean;
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
  if (!Number.isInteger(value) || value < min) throw new ConfigError(`${name} must be an integer >= ${min}, got "${raw}"`);
  return value;
}

function bool(env: Env, name: string, fallback: boolean): boolean {
  const raw = env[name]?.trim().toLowerCase();
  if (!raw) return fallback;
  return raw === 'true' || raw === '1' || raw === 'yes';
}

function required(env: Env, name: string, minLength: number): string {
  const value = env[name]?.trim();
  if (!value) throw new ConfigError(`${name} is required`);
  if (value.length < minLength) throw new ConfigError(`${name} must be at least ${minLength} characters`);
  return value;
}

function readPem(path: string, name: string): Buffer {
  if (!existsSync(path)) {
    throw new ConfigError(`${name} not found at ${path}. Run "npm run certs" from the repo root, or set ${name}.`);
  }
  return readFileSync(path);
}

export function loadConfig(env: Env = process.env, cwd: string = process.cwd()): DemoConfig {
  const tls = bool(env, 'TLS_DISABLED', false)
    ? null
    : {
        cert: readPem(resolve(cwd, str(env, 'TLS_CERT', '../../certs/localhost.pem')), 'TLS_CERT'),
        key: readPem(resolve(cwd, str(env, 'TLS_KEY', '../../certs/localhost-key.pem')), 'TLS_KEY'),
      };
  return {
    host: str(env, 'HOST', 'localhost'),
    port: int(env, 'PORT', 3443, 0),
    tls,
    dataDir: resolve(cwd, str(env, 'DATA_DIR', 'data')),
    gatewaySecret: required(env, 'GATEWAY_SECRET', 16),
    cookieSecret: required(env, 'COOKIE_SECRET', 32),
    davPublicUrl: str(env, 'DAV_PUBLIC_URL', 'https://localhost:8443').replace(/\/+$/, ''),
    tokenMaxAgeS: int(env, 'TOKEN_MAX_AGE_S', 7 * 24 * 3600, 1),
    lockMaxTimeoutS: int(env, 'LOCK_MAX_TIMEOUT_S', 3600, 1),
    maxUploadBytes: int(env, 'MAX_UPLOAD_MB', 512, 1) * 1024 * 1024,
    seed: bool(env, 'SEED', true),
    logLevel: str(env, 'LOG_LEVEL', 'info'),
  };
}
