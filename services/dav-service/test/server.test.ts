import { Writable } from 'node:stream';
import type { FastifyInstance } from 'fastify';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { DavConfig } from '../src/config.js';
import { GATEWAY_UNAVAILABLE, GatewayError } from '../src/gateway/errors.js';
import { createLogger } from '../src/logging.js';
import { buildServer } from '../src/server.js';
import { FakeGateway } from './fake-gateway.js';

const config: DavConfig = {
  host: '127.0.0.1',
  port: 0,
  tls: null,
  publicBaseUrl: 'https://dav.example',
  gatewayUrl: 'http://unused',
  gatewaySecret: 'unused',
  gatewayCa: undefined,
  maxUploadBytes: 1024 * 1024,
  lockMaxTimeoutS: 600,
  resolveCacheMs: 5000,
  logHeaders: true,
  logLevel: 'info',
};

const LOCKINFO = `<?xml version="1.0" encoding="utf-8"?>
<D:lockinfo xmlns:D="DAV:"><D:lockscope><D:exclusive/></D:lockscope><D:locktype><D:write/></D:locktype>
<D:owner><D:href>CORP\\alice</D:href></D:owner></D:lockinfo>`;

let gateway: FakeGateway;
let app: FastifyInstance;
let logs: string[];

beforeEach(async () => {
  gateway = new FakeGateway();
  logs = [];
  const sink = new Writable({
    write(chunk, _enc, cb) {
      logs.push(String(chunk));
      cb();
    },
  });
  app = buildServer(config, { gateway, logger: createLogger('debug', sink) });
  await app.ready();
});

afterEach(async () => {
  await app.close();
});

function setup(name = 'Plan.docx', opts = {}) {
  const file = gateway.addFile(name, 'original');
  const { token, session } = gateway.addSession(file.fileId, opts);
  const collection = `/s/${token}/${file.fileId}/`;
  const path = `${collection}${encodeURIComponent(name)}`;
  return { file, token, session, collection, path };
}

async function lock(path: string) {
  const res = await app.inject({ method: 'LOCK', url: path, headers: { timeout: 'Second-3600' }, payload: LOCKINFO });
  expect(res.statusCode).toBe(200);
  return /<(.+)>/.exec(res.headers['lock-token'] as string)![1]!;
}

describe('OPTIONS and probes', () => {
  it('answers OPTIONS on / without touching the gateway', async () => {
    const res = await app.inject({ method: 'OPTIONS', url: '/' });
    expect(res.statusCode).toBe(200);
    expect(res.headers.dav).toBe('1,2');
    expect(res.headers['ms-author-via']).toBe('DAV');
    expect(res.headers.allow).toBe('OPTIONS, GET, HEAD, PUT, PROPFIND, PROPPATCH, LOCK, UNLOCK');
    expect(gateway.calls).toHaveLength(0);
  });

  it('answers OPTIONS on the file after resolving the token', async () => {
    const { path } = setup();
    const res = await app.inject({ method: 'OPTIONS', url: path });
    expect(res.statusCode).toBe(200);
    expect(res.headers.dav).toBe('1,2');
    expect(gateway.calls[0]!.op).toBe('resolve');
  });

  it('gives SharePoint and FrontPage probes a fast 404', async () => {
    for (const url of ['/_vti_inf.html', '/_vti_bin/shtml.dll/_vti_rpc', '/s/x/y/_vti_bin/owssvr.dll']) {
      const res = await app.inject({ method: 'POST', url });
      expect(res.statusCode).toBe(404);
    }
    expect(gateway.calls).toHaveLength(0);
  });

  it('forbids MOVE, COPY, DELETE and MKCOL', async () => {
    const { path, collection } = setup();
    for (const method of ['MOVE', 'COPY', 'DELETE'] as const) {
      expect((await app.inject({ method, url: path })).statusCode).toBe(403);
    }
    expect((await app.inject({ method: 'MKCOL', url: `${collection}new` })).statusCode).toBe(403);
  });
});

describe('PROPFIND', () => {
  it('answers Depth 0 on the root and session collections', async () => {
    const { token } = setup();
    for (const url of ['/', `/s/${token}/`]) {
      const res = await app.inject({ method: 'PROPFIND', url, headers: { depth: '1' } });
      expect(res.statusCode).toBe(207);
      expect(res.body.match(/<D:response>/g)).toHaveLength(1);
      expect(res.body).toContain('<D:collection/>');
    }
  });

  it('returns the file at Depth 0', async () => {
    const { path, file } = setup();
    const res = await app.inject({ method: 'PROPFIND', url: path, headers: { depth: '0' } });
    expect(res.statusCode).toBe(207);
    expect(res.headers['content-type']).toMatch(/^application\/xml/);
    expect(res.body).toContain(`<D:getetag>&quot;${file.versionId}&quot;</D:getetag>`);
    expect(res.body).toContain(`<D:getcontentlength>${file.size}</D:getcontentlength>`);
    expect(res.body).toContain('<D:resourcetype/>');
  });

  it('lists the collection and the file at Depth 1', async () => {
    const { collection, file, token } = setup('תוכנית עבודה (טיוטה).docx');
    const res = await app.inject({ method: 'PROPFIND', url: collection, headers: { depth: '1' }, payload: '' });
    expect(res.statusCode).toBe(207);
    const hrefs = [...res.body.matchAll(/<D:href>([^<]+)<\/D:href>/g)].map((m) => m[1]!);
    expect(hrefs).toEqual([
      `https://dav.example/s/${token}/${file.fileId}/`,
      `https://dav.example/s/${token}/${file.fileId}/%D7%AA%D7%95%D7%9B%D7%A0%D7%99%D7%AA%20%D7%A2%D7%91%D7%95%D7%93%D7%94%20%28%D7%98%D7%99%D7%95%D7%98%D7%94%29.docx`,
    ]);
    expect(decodeURIComponent(hrefs[1]!.split('/').pop()!)).toBe('תוכנית עבודה (טיוטה).docx');
  });

  it('reports unknown properties as 404 propstat', async () => {
    const { path } = setup();
    const res = await app.inject({
      method: 'PROPFIND',
      url: path,
      headers: { depth: '0', 'content-type': 'text/xml; charset="utf-8"' },
      payload:
        '<D:propfind xmlns:D="DAV:" xmlns:Z="urn:schemas-microsoft-com:"><D:prop><D:displayname/><Z:Win32FileAttributes/></D:prop></D:propfind>',
    });
    expect(res.statusCode).toBe(207);
    expect(res.body).toContain('<D:displayname>Plan.docx</D:displayname>');
    expect(res.body).toMatch(/Win32FileAttributes[^]*404 Not Found/);
  });

  it('rejects malformed XML with 400', async () => {
    const { path } = setup();
    const res = await app.inject({ method: 'PROPFIND', url: path, payload: '<D:propfind xmlns:D="DAV:">' });
    expect(res.statusCode).toBe(400);
  });

  it('returns 404 for a fileId outside the session, or another name', async () => {
    const { token, file } = setup();
    expect((await app.inject({ method: 'PROPFIND', url: `/s/${token}/other-id/Plan.docx` })).statusCode).toBe(404);
    expect((await app.inject({ method: 'PROPFIND', url: `/s/${token}/${file.fileId}/~$Plan.docx` })).statusCode).toBe(404);
  });

  it('answers PROPPATCH with success without storing anything', async () => {
    const { path } = setup();
    const res = await app.inject({
      method: 'PROPPATCH',
      url: path,
      payload:
        '<D:propertyupdate xmlns:D="DAV:" xmlns:Z="urn:schemas-microsoft-com:"><D:set><D:prop><Z:Win32LastAccessTime>x</Z:Win32LastAccessTime></D:prop></D:set></D:propertyupdate>',
    });
    expect(res.statusCode).toBe(207);
    expect(res.body).toMatch(/Win32LastAccessTime[^]*200 OK/);
  });
});

describe('GET and HEAD', () => {
  it('streams the content with exact headers', async () => {
    const { path, file } = setup();
    const res = await app.inject({ method: 'GET', url: path });
    expect(res.statusCode).toBe(200);
    expect(res.body).toBe('original');
    expect(res.headers.etag).toBe(`"${file.versionId}"`);
    expect(res.headers['content-type']).toBe(file.mime);
    expect(res.headers['content-length']).toBe('8');
    expect(res.headers['last-modified']).toBe('Fri, 02 Jan 2026 10:00:00 GMT');
  });

  it('answers HEAD from metadata alone', async () => {
    const { path, file } = setup();
    const res = await app.inject({ method: 'HEAD', url: path });
    expect(res.statusCode).toBe(200);
    expect(res.headers.etag).toBe(`"${file.versionId}"`);
    expect(res.headers['content-length']).toBe('8');
    expect(gateway.calls.map((c) => c.op)).not.toContain('getContent');
  });

  it('honours If-None-Match', async () => {
    const { path, file } = setup();
    const res = await app.inject({ method: 'GET', url: path, headers: { 'if-none-match': `"${file.versionId}"` } });
    expect(res.statusCode).toBe(304);
  });
});

describe('LOCK, PUT and UNLOCK', () => {
  it('runs the Office save sequence', async () => {
    const { path, file, session } = setup();
    const lockRes = await app.inject({
      method: 'LOCK',
      url: path,
      headers: { timeout: 'Second-3600', depth: '0' },
      payload: LOCKINFO,
    });
    expect(lockRes.statusCode).toBe(200);
    const lockToken = /^<(opaquelocktoken:[^>]+)>$/.exec(lockRes.headers['lock-token'] as string)![1]!;
    expect(lockRes.body).toContain('<D:owner xmlns:D="DAV:"><D:href>CORP\\alice</D:href></D:owner>');
    // Timeout is capped at LOCK_MAX_TIMEOUT_S.
    expect(lockRes.body).toContain('<D:timeout>Second-600</D:timeout>');
    const lockCall = gateway.calls.find((c) => c.op === 'lock')!;
    expect(lockCall.arg).toMatchObject({ fileId: file.fileId, sessionId: session.sessionId, timeoutS: 600 });
    expect(lockCall.ctx.sessionId).toBe(session.sessionId);
    expect(lockCall.ctx.requestId).toBeTruthy();

    const putRes = await app.inject({
      method: 'PUT',
      url: path,
      headers: { if: `<https://dav.example${path}> (<${lockToken}>)` },
      payload: 'edited',
    });
    expect(putRes.statusCode).toBe(204);
    const stored = gateway.files.get(file.fileId)!;
    expect(stored.content.toString()).toBe('edited');
    expect(putRes.headers.etag).toBe(`"${stored.info.versionId}"`);

    const refresh = await app.inject({ method: 'LOCK', url: path, headers: { if: `(<${lockToken}>)`, timeout: 'Second-60' } });
    expect(refresh.statusCode).toBe(200);
    expect(refresh.headers['lock-token']).toBeUndefined();
    expect(refresh.body).toContain('<D:timeout>Second-60</D:timeout>');
    expect(refresh.body).toContain('CORP\\alice');

    const unlock = await app.inject({ method: 'UNLOCK', url: path, headers: { 'lock-token': `<${lockToken}>` } });
    expect(unlock.statusCode).toBe(204);
    expect(stored.info.lock).toBeNull();
  });

  it('shows the lock in lockdiscovery', async () => {
    const { path } = setup();
    const lockToken = await lock(path);
    const res = await app.inject({
      method: 'PROPFIND',
      url: path,
      headers: { depth: '0' },
      payload: '<D:propfind xmlns:D="DAV:"><D:prop><D:lockdiscovery/></D:prop></D:propfind>',
    });
    expect(res.body).toContain(`<D:locktoken><D:href>${lockToken}</D:href></D:locktoken>`);
    expect(res.body).toMatch(/<D:timeout>Second-\d+<\/D:timeout>/);
  });

  it('refuses a PUT without a lock token before calling the gateway', async () => {
    const { path } = setup();
    const res = await app.inject({ method: 'PUT', url: path, payload: 'x' });
    expect(res.statusCode).toBe(423);
    expect(gateway.calls.map((c) => c.op)).not.toContain('putContent');
  });

  it('maps a gateway 412 on PUT', async () => {
    const { path } = setup();
    await lock(path);
    const res = await app.inject({ method: 'PUT', url: path, headers: { if: '(<opaquelocktoken:stale>)' }, payload: 'x' });
    expect(res.statusCode).toBe(412);
  });

  it('refuses shared locks with 412', async () => {
    const { path } = setup();
    const res = await app.inject({
      method: 'LOCK',
      url: path,
      payload:
        '<D:lockinfo xmlns:D="DAV:"><D:lockscope><D:shared/></D:lockscope><D:locktype><D:write/></D:locktype></D:lockinfo>',
    });
    expect(res.statusCode).toBe(412);
  });

  it('maps a lock conflict to 423', async () => {
    const { path } = setup();
    await lock(path);
    const res = await app.inject({ method: 'LOCK', url: path, payload: LOCKINFO });
    expect(res.statusCode).toBe(423);
  });

  it('requires a Lock-Token header on UNLOCK', async () => {
    const { path } = setup();
    expect((await app.inject({ method: 'UNLOCK', url: path })).statusCode).toBe(400);
  });

  it('enforces read-only sessions', async () => {
    const { path } = setup('Plan.docx', { canWrite: false });
    expect((await app.inject({ method: 'LOCK', url: path, payload: LOCKINFO })).statusCode).toBe(403);
    expect((await app.inject({ method: 'PUT', url: path, headers: { if: '(<opaquelocktoken:x>)' }, payload: 'x' })).statusCode).toBe(403);
  });

  it('refuses uploads above MAX_UPLOAD_MB', async () => {
    const { path } = setup();
    const lockToken = await lock(path);
    const res = await app.inject({
      method: 'PUT',
      url: path,
      headers: { if: `(<${lockToken}>)` },
      payload: Buffer.alloc(config.maxUploadBytes + 1),
    });
    expect(res.statusCode).toBe(413);
  });
});

describe('gateway failures', () => {
  it('returns 401 for an unknown token', async () => {
    const res = await app.inject({ method: 'PROPFIND', url: '/s/nope/f1/a.docx' });
    expect(res.statusCode).toBe(401);
    expect(res.headers['www-authenticate']).toMatch(/^Bearer/);
  });

  it('caches resolve, and drops the entry after a 401', async () => {
    const { path, token } = setup();
    await app.inject({ method: 'HEAD', url: path });
    await app.inject({ method: 'HEAD', url: path });
    expect(gateway.calls.filter((c) => c.op === 'resolve')).toHaveLength(1);

    // The gateway revokes the session: it now fails file calls with 401.
    gateway.failWith = new GatewayError(401, { error: 'session_invalid', reason: 'revoked' }, 'revoked');
    expect((await app.inject({ method: 'HEAD', url: path })).statusCode).toBe(401);
    gateway.failWith = null;
    gateway.sessions.delete(token);
    expect((await app.inject({ method: 'HEAD', url: path })).statusCode).toBe(401);
  });

  it('returns 503 when the gateway is unreachable', async () => {
    const { path } = setup();
    gateway.failWith = new GatewayError(GATEWAY_UNAVAILABLE, null, 'ECONNREFUSED');
    expect((await app.inject({ method: 'GET', url: path })).statusCode).toBe(503);
  });
});

describe('logging', () => {
  it('never logs the token, and names the session instead', async () => {
    const { path, token, session } = setup();
    const lockToken = await lock(path);
    await app.inject({ method: 'PUT', url: path, headers: { if: `<https://dav.example${path}> (<${lockToken}>)` }, payload: 'x' });
    await app.inject({ method: 'PROPFIND', url: `/s/${token}/wrong/x.docx` });
    const all = logs.join('');
    expect(all).not.toContain(token);
    expect(all).toContain('/s/[token]/');
    expect(all).toContain(session.sessionId);
    // DAV_LOG_HEADERS=true logs request and response headers.
    expect(all).toContain('"responseHeaders"');
    expect(all).toContain('"headers"');
  });
});
