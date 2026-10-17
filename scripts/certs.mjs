#!/usr/bin/env node
// Creates the localhost TLS certificate with mkcert and installs its root CA.
// Office requires HTTPS with a certificate that Windows trusts.
import { execFileSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const certs = join(root, 'certs');

function mkcert(args, opts = {}) {
  return execFileSync('mkcert', args, { cwd: root, encoding: 'utf8', stdio: opts.capture ? 'pipe' : 'inherit' });
}

try {
  mkcert(['-version'], { capture: true });
} catch {
  console.error(`mkcert is not installed.

  Windows:  winget install FiloSottile.mkcert   (or: choco install mkcert / scoop install mkcert)
  macOS:    brew install mkcert
  Linux:    see https://github.com/FiloSottile/mkcert#linux

Then run "npm run certs" again.`);
  process.exit(1);
}

mkdirSync(certs, { recursive: true });
mkcert(['-install']);
mkcert([
  '-cert-file',
  join(certs, 'localhost.pem'),
  '-key-file',
  join(certs, 'localhost-key.pem'),
  'localhost',
  '127.0.0.1',
  '::1',
]);

// dav-service calls the gateway over HTTPS and needs to trust the same root.
const caroot = mkcert(['-CAROOT'], { capture: true }).trim();
const rootCa = join(caroot, 'rootCA.pem');
if (existsSync(rootCa)) {
  copyFileSync(rootCa, join(certs, 'rootCA.pem'));
  console.log(`Copied ${rootCa} to certs/rootCA.pem (used as GATEWAY_CA_FILE).`);
}

const isWsl = existsSync('/proc/version') && /microsoft/i.test(readFileSync('/proc/version', 'utf8'));
if (isWsl) {
  console.log(`
You are in WSL. Office trusts the Windows certificate store, not WSL's.
Make Windows trust this same root CA, for example from PowerShell:

  certutil -user -addstore Root "\\\\wsl.localhost\\${process.env.WSL_DISTRO_NAME ?? '<distro>'}${join(certs, 'rootCA.pem').replace(/\//g, '\\')}"

or share one CA between both sides: run "mkcert -install" on Windows, then in WSL
  export CAROOT=/mnt/c/Users/<you>/AppData/Local/mkcert
and run "npm run certs" again. See the README.`);
}
