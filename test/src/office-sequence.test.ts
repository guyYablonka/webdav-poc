import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { LOCKINFO, PROPFIND_OFFICE, davClient, demoClient, header, hrefs, openEvents, sleep } from './clients.js';
import { startStack, type Stack } from './stack.js';

/** Short enough to expire inside a test, long enough not to expire between two requests. */
const LOCK_MAX_TIMEOUT_S = 3;
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

let stack: Stack;
let dav: ReturnType<typeof davClient>;
let demo: ReturnType<typeof demoClient>;
let alice: string;
let bob: string;
const tokensSeen: string[] = [];

beforeAll(async () => {
  stack = await startStack({ lockMaxTimeoutS: LOCK_MAX_TIMEOUT_S, resolveCacheMs: 5000 });
  dav = davClient(stack);
  demo = demoClient(stack);
  alice = await demo.login('alice');
  bob = await demo.login('bob');
});

afterAll(async () => {
  await stack?.stop();
});

/** A fresh file per test keeps lock state independent. */
async function newFile(name: string): Promise<string> {
  return demo.upload(alice, name, Buffer.from(`PK fake content of ${name}`));
}

async function openForEdit(cookie: string, fileId: string) {
  const session = await demo.editSession(cookie, fileId);
  tokensSeen.push(session.token);
  return session;
}

async function lock(url: string, extra: Record<string, string> = {}) {
  const res = await dav('LOCK', url, {
    headers: { 'content-type': 'text/xml; charset="utf-8"', timeout: 'Second-3600', depth: '0', ...extra },
    body: LOCKINFO,
  });
  const token = /^<(.+)>$/.exec(header(res, 'lock-token') ?? '')?.[1];
  return { res, token };
}

describe('Office request sequence', () => {
  it('opens, saves twice, refreshes the lock and closes', async () => {
    const files = await demo.files(alice);
    const plan = files.find((f) => f.name === 'Project Plan.docx')!;
    const { davUrl, uri } = await openForEdit(alice, plan.id);
    expect(uri).toBe(`ms-word:ofe|u|${davUrl}`);
    const collection = davUrl.slice(0, davUrl.lastIndexOf('/') + 1);
    const events = await openEvents(stack, alice);

    // 1. OPTIONS on the root, then on the file.
    let res = await dav('OPTIONS', `${stack.davUrl}/`, { headers: { 'user-agent': 'Microsoft Office Protocol Discovery' } });
    expect(res.status).toBe(200);
    expect(header(res, 'dav')).toBe('1,2');
    expect(header(res, 'ms-author-via')).toBe('DAV');
    expect(header(res, 'allow')).toBe('OPTIONS, GET, HEAD, PUT, PROPFIND, PROPPATCH, LOCK, UNLOCK');
    res = await dav('OPTIONS', davUrl);
    expect(res.status).toBe(200);
    expect(header(res, 'dav')).toBe('1,2');

    // 2. PROPFIND Depth 0 on the file, then Depth 1 on its collection.
    res = await dav('PROPFIND', davUrl, { headers: { depth: '0', 'content-type': 'text/xml' }, body: PROPFIND_OFFICE });
    expect(res.status).toBe(207);
    const detail = await demo.file(alice, plan.id);
    const v1 = detail.versions[0]!.id;
    expect(res.text).toContain(`<D:getetag>&quot;${v1}&quot;</D:getetag>`);
    expect(res.text).toContain(`<D:getcontenttype>${DOCX}</D:getcontenttype>`);
    expect(res.text).toMatch(/<D:getlastmodified>\w{3}, \d{2} \w{3} \d{4} \d{2}:\d{2}:\d{2} GMT<\/D:getlastmodified>/);
    expect(res.text).toMatch(/Win32FileAttributes[^]*404 Not Found/);
    res = await dav('PROPFIND', collection, { headers: { depth: '1' } });
    expect(res.status).toBe(207);
    expect(hrefs(res.text)).toEqual([collection, davUrl]);

    // 3. LOCK, then GET.
    const locked = await lock(davUrl);
    expect(locked.res.status).toBe(200);
    expect(locked.token).toMatch(/^opaquelocktoken:/);
    expect(locked.res.text).toContain(`<D:timeout>Second-${LOCK_MAX_TIMEOUT_S}</D:timeout>`);
    expect(locked.res.text).toContain('<D:owner xmlns:D="DAV:"><D:href>CONTOSO\\alice</D:href></D:owner>');
    const lockToken = locked.token!;
    await events.waitFor((e) => e.type === 'lock.acquired' && e.fileId === plan.id);

    res = await dav('GET', davUrl);
    expect(res.status).toBe(200);
    expect(header(res, 'etag')).toBe(`"${v1}"`);
    expect(header(res, 'content-type')).toBe(DOCX);
    expect(header(res, 'content-length')).toBe(String(res.body.length));
    expect(header(res, 'last-modified')).toMatch(/GMT$/);
    expect(res.body.subarray(0, 2).toString()).toBe('PK');
    const original = res.body;

    // 4. PUT with If: (<token>), then check the ETag and the new version.
    res = await dav('PUT', davUrl, { headers: { if: `(<${lockToken}>)`, 'content-type': DOCX }, body: original });
    expect(res.status).toBe(204);
    const etag2 = header(res, 'etag')!;
    expect(etag2).toMatch(/^"[0-9a-f-]{36}"$/);
    expect(etag2).not.toBe(`"${v1}"`);
    const created = await events.waitFor((e) => e.type === 'version.created' && e.fileId === plan.id);
    expect(`"${created.versionId}"`).toBe(etag2);
    let after = await demo.file(alice, plan.id);
    expect(after.versionCount).toBe(2);
    expect(after.versions[0]).toMatchObject({ versionNo: 2, source: 'dav', createdBy: { id: 'alice' } });
    expect(`"${after.versions[0]!.id}"`).toBe(etag2);
    res = await dav('GET', davUrl);
    expect(header(res, 'etag')).toBe(etag2);
    expect(res.body.equals(original)).toBe(true);

    // 5. LOCK refresh, a second PUT (tagged If header), then UNLOCK.
    res = await dav('LOCK', davUrl, { headers: { if: `(<${lockToken}>)`, timeout: 'Second-3600' } });
    expect(res.status).toBe(200);
    expect(res.text).toContain(lockToken);
    res = await dav('PUT', davUrl, { headers: { if: `<${davUrl}> (<${lockToken}>)` }, body: Buffer.concat([original, Buffer.from('\n')]) });
    expect(res.status).toBe(204);
    const etag3 = header(res, 'etag')!;
    expect(etag3).not.toBe(etag2);
    res = await dav('UNLOCK', davUrl, { headers: { 'lock-token': `<${lockToken}>` } });
    expect(res.status).toBe(204);
    await events.waitFor((e) => e.type === 'lock.released' && e.fileId === plan.id);

    after = await demo.file(alice, plan.id);
    expect(after.versionCount).toBe(3);
    expect(after.lock).toBeNull();
    res = await dav('PROPFIND', davUrl, { headers: { depth: '0' }, body: PROPFIND_OFFICE });
    expect(res.text).toContain('<D:lockdiscovery/>');
    expect(res.text).toContain(`<D:getetag>${etag3.replace(/"/g, '&quot;')}</D:getetag>`);

    expect(events.events.filter((e) => e.fileId === plan.id).map((e) => e.type)).toEqual([
      'lock.acquired',
      'version.created',
      'lock.refreshed',
      'version.created',
      'lock.released',
    ]);
    events.close();
  });

  it('answers the parent-path probes and SharePoint probes', async () => {
    const fileId = await newFile('Probe.docx');
    const { token } = await openForEdit(alice, fileId);
    for (const url of [`${stack.davUrl}/`, `${stack.davUrl}/s/${token}/`]) {
      const res = await dav('PROPFIND', url, { headers: { depth: '0' } });
      expect(res.status).toBe(207);
      expect(res.text).toContain('<D:collection/>');
    }
    for (const path of ['/_vti_inf.html', '/_vti_bin/shtml.dll/_vti_rpc', `/s/${token}/${fileId}/_vti_bin/owssvr.dll`]) {
      expect((await dav('POST', `${stack.davUrl}${path}`)).status).toBe(404);
    }
  });
});

describe('failure paths', () => {
  it('a second session gets 423 on LOCK', async () => {
    const fileId = await newFile('Contested.docx');
    const a = await openForEdit(alice, fileId);
    const b = await openForEdit(bob, fileId);
    expect((await lock(a.davUrl)).res.status).toBe(200);
    const second = await lock(b.davUrl);
    expect(second.res.status).toBe(423);
    // Bob's Office reads the holder from lockdiscovery to say who is editing.
    const res = await dav('PROPFIND', b.davUrl, { headers: { depth: '0' }, body: PROPFIND_OFFICE });
    expect(res.text).toContain('CONTOSO\\alice');
    expect((await demo.files(alice)).find((f) => f.id === fileId)!.lock?.user.id).toBe('alice');
  });

  it('a PUT without a lock gets 423', async () => {
    const fileId = await newFile('NoLock.docx');
    const { davUrl } = await openForEdit(alice, fileId);
    expect((await dav('PUT', davUrl, { body: 'x' })).status).toBe(423);
    // ...and a lock token that is not the file's lock is refused too.
    expect((await dav('PUT', davUrl, { headers: { if: '(<opaquelocktoken:bogus>)' }, body: 'x' })).status).toBe(412);
    expect((await demo.file(alice, fileId)).versionCount).toBe(1);
  });

  it('a revoked session gets 401, and a save fails cleanly', async () => {
    const fileId = await newFile('Revoked.docx');
    const { davUrl, sessionId } = await openForEdit(alice, fileId);
    const { token: lockToken } = await lock(davUrl);
    // Warm dav-service's resolve cache, so the test also shows revocation beats the cache.
    expect((await dav('HEAD', davUrl)).status).toBe(200);
    expect(await demo.revoke(alice, sessionId)).toBe(204);

    const put = await dav('PUT', davUrl, { headers: { if: `(<${lockToken}>)` }, body: 'after revoke' });
    expect(put.status).toBe(401);
    expect(header(put, 'www-authenticate')).toMatch(/^Bearer /);
    expect((await dav('PROPFIND', davUrl, { headers: { depth: '0' } })).status).toBe(401);
    expect((await dav('GET', davUrl)).status).toBe(401);
    expect((await dav('LOCK', davUrl, { headers: { if: `(<${lockToken}>)` } })).status).toBe(401);
    const detail = await demo.file(alice, fileId);
    expect(detail.versionCount).toBe(1);
    expect(detail.lock).toBeNull();
  });

  it('logging out revokes the user’s DAV sessions', async () => {
    const fileId = await newFile('Logout.docx');
    const carol = await demo.login('bob');
    const { davUrl } = await openForEdit(carol, fileId);
    expect((await dav('HEAD', davUrl)).status).toBe(200);
    await demo.logout(carol);
    expect((await dav('GET', davUrl)).status).toBe(401);
  });

  it('an expired lock on an unchanged file is revived by PUT', async () => {
    const fileId = await newFile('Sleepy.docx');
    const { davUrl } = await openForEdit(alice, fileId);
    const events = await openEvents(stack, alice);
    const { token: lockToken } = await lock(davUrl);
    await sleep(LOCK_MAX_TIMEOUT_S * 1000 + 500);
    // The laptop slept through the lock timeout: nobody else touched the file.
    expect((await demo.files(alice)).find((f) => f.id === fileId)!.lock).toBeNull();

    const res = await dav('PUT', davUrl, { headers: { if: `(<${lockToken}>)` }, body: 'saved after sleep' });
    expect(res.status).toBe(204);
    const revived = await events.waitFor((e) => e.type === 'lock.revived' && e.fileId === fileId);
    expect(revived.userId).toBe('alice');
    expect(stack.logs.demo.join('\n')).toContain('lock revived by PUT');
    expect((await demo.file(alice, fileId)).versionCount).toBe(2);
    // The revived lock is active again and blocks others.
    expect((await demo.files(alice)).find((f) => f.id === fileId)!.lock?.user.id).toBe('alice');
    events.close();
  });

  it('an expired lock where another session saved in between gets 412', async () => {
    const fileId = await newFile('Overtaken.docx');
    const a = await openForEdit(alice, fileId);
    const { token: aliceLock } = await lock(a.davUrl);
    await sleep(LOCK_MAX_TIMEOUT_S * 1000 + 500);

    const b = await openForEdit(bob, fileId);
    const { token: bobLock } = await lock(b.davUrl);
    expect(bobLock).toBeTruthy();
    expect((await dav('PUT', b.davUrl, { headers: { if: `(<${bobLock}>)` }, body: 'bob wins' })).status).toBe(204);
    expect((await dav('UNLOCK', b.davUrl, { headers: { 'lock-token': `<${bobLock}>` } })).status).toBe(204);

    const res = await dav('PUT', a.davUrl, { headers: { if: `(<${aliceLock}>)` }, body: 'alice is late' });
    expect(res.status).toBe(412);
    const detail = await demo.file(alice, fileId);
    expect(detail.versionCount).toBe(2);
    expect(detail.versions[0]!.createdBy.id).toBe('bob');
  });

  it('force unlock makes the next save fail with 412', async () => {
    const fileId = await newFile('Forced.docx');
    const { davUrl } = await openForEdit(alice, fileId);
    const { token: lockToken } = await lock(davUrl);
    expect(await demo.forceUnlock(bob, fileId)).toBe(204);
    expect((await dav('PUT', davUrl, { headers: { if: `(<${lockToken}>)` }, body: 'x' })).status).toBe(412);
  });

  it('a web upload while locked gets 409', async () => {
    const fileId = await newFile('WebUpload.docx');
    const { davUrl } = await openForEdit(alice, fileId);
    await lock(davUrl);
    expect(await demo.uploadVersion(bob, fileId, 'WebUpload.docx', Buffer.from('web'))).toBe(409);
  });

  it('a wrong fileId in the path gets 404', async () => {
    const fileId = await newFile('Mine.docx');
    const otherId = await newFile('NotMine.docx');
    const { token } = await openForEdit(alice, fileId);
    expect((await dav('PROPFIND', `${stack.davUrl}/s/${token}/${otherId}/NotMine.docx`, { headers: { depth: '0' } })).status).toBe(404);
    expect((await dav('GET', `${stack.davUrl}/s/${token}/${otherId}/NotMine.docx`)).status).toBe(404);
    expect((await dav('PROPFIND', `${stack.davUrl}/s/${token}/${otherId}/`, { headers: { depth: '1' } })).status).toBe(404);
  });

  it('a Hebrew file name with spaces and parentheses round-trips in hrefs', async () => {
    const name = 'תוכנית עבודה (טיוטה).docx';
    const seeded = (await demo.files(alice)).find((f) => f.name === name)!;
    const { davUrl } = await openForEdit(alice, seeded.id);
    expect(davUrl).not.toMatch(/[ ()֐-׿]/);
    const collection = davUrl.slice(0, davUrl.lastIndexOf('/') + 1);

    const res = await dav('PROPFIND', collection, { headers: { depth: '1' }, body: PROPFIND_OFFICE });
    expect(res.status).toBe(207);
    const [collectionHref, fileHref] = hrefs(res.text);
    expect(collectionHref).toBe(collection);
    expect(fileHref).toBe(davUrl);
    expect(decodeURIComponent(new URL(fileHref!).pathname.split('/').pop()!)).toBe(name);
    expect(res.text).toContain(`<D:displayname>${name}</D:displayname>`);

    // The href is directly usable, and so is the same name encoded differently.
    expect((await dav('GET', fileHref!)).status).toBe(200);
    const lenient = `${collection}${encodeURIComponent(name)}`;
    expect(lenient).not.toBe(fileHref);
    expect((await dav('HEAD', lenient)).status).toBe(200);
  });

  it('streams a large PUT, and refuses one with the wrong token promptly', async () => {
    const fileId = await newFile('Large.xlsx');
    const { davUrl } = await openForEdit(alice, fileId);
    const { token: lockToken } = await lock(davUrl);
    const big = Buffer.alloc(20 * 1024 * 1024, 7);
    let res = await dav('PUT', davUrl, { headers: { if: `(<${lockToken}>)` }, body: big });
    expect(res.status).toBe(204);
    expect((await demo.file(alice, fileId)).versions[0]).toMatchObject({ versionNo: 2 });
    res = await dav('GET', davUrl);
    expect(res.body.length).toBe(big.length);

    // The gateway refuses before reading the body (the file is locked under another token).
    const started = Date.now();
    res = await dav('PUT', davUrl, { headers: { if: '(<opaquelocktoken:stale>)' }, body: big });
    expect(res.status).toBe(423);
    expect(Date.now() - started).toBeLessThan(10_000);
    // The connection survives: the next request still works.
    expect((await dav('HEAD', davUrl)).status).toBe(200);
  });
});

describe('logging', () => {
  it('never writes a session token to either service’s logs', () => {
    expect(tokensSeen.length).toBeGreaterThan(5);
    const davLogs = stack.logs.dav.join('\n');
    const demoLogs = stack.logs.demo.join('\n');
    for (const token of tokensSeen) {
      expect(davLogs).not.toContain(token);
      expect(demoLogs).not.toContain(token);
    }
    // DAV_LOG_HEADERS=true is on: headers are logged, with the token redacted.
    expect(davLogs).toContain('"dav request"');
    expect(davLogs).toContain('/s/[token]/');
    expect(davLogs).toMatch(/"sessionId":"[0-9a-f-]{36}"/);
  });
});
