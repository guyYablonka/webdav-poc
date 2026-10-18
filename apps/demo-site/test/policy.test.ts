import { existsSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { Readable } from 'node:stream';
import { describe, expect, it } from 'vitest';
import { hashToken } from '../src/domain/dav-policy.js';
import { PolicyError } from '../src/domain/errors.js';
import { LOCK_MAX_TIMEOUT_S, MAX_UPLOAD_BYTES, TOKEN_MAX_AGE_S, body, createWorld } from './helpers.js';

/** Runs `fn` and returns the PolicyError it throws (sync or async). */
async function refusal(fn: () => unknown): Promise<PolicyError> {
  try {
    await fn();
  } catch (err) {
    if (err instanceof PolicyError) return err;
    throw err;
  }
  throw new Error('expected a PolicyError');
}

describe('sessions', () => {
  it('issues 32-byte base64url tokens and stores only their hash', async () => {
    const w = await createWorld();
    const { token, session } = w.session();
    expect(token).toMatch(/^[A-Za-z0-9_-]{43}$/);
    expect(session.token_hash).toBe(hashToken(token));
    const dump = JSON.stringify(w.repo.db.prepare('SELECT * FROM dav_sessions').all());
    expect(dump).not.toContain(token);
    expect(session.max_expires_at - session.created_at).toBe(TOKEN_MAX_AGE_S * 1000);
  });

  it('resolves an active session', async () => {
    const w = await createWorld();
    const { token, session } = w.session();
    expect(w.policy.resolve(token, 'read')).toEqual({
      sessionId: session.id,
      userId: 'alice',
      displayName: 'Alice',
      fileId: w.file.id,
      canWrite: true,
      state: 'active',
    });
  });

  it('is cheap and idempotent: resolve changes nothing', async () => {
    const w = await createWorld();
    const { token, session } = w.session();
    w.advance(10);
    w.policy.resolve(token, 'write');
    w.policy.resolve(token, 'write');
    expect(w.repo.getSession(session.id)!.last_seen_at).toBe(session.last_seen_at);
  });

  it('rejects unknown tokens', async () => {
    const w = await createWorld();
    const err = await refusal(() => w.policy.resolve('nope', 'read'));
    expect([err.status, err.details.reason]).toEqual([401, 'unknown']);
  });

  it('has no idle timeout: an untouched session stays valid until the cap', async () => {
    const w = await createWorld();
    const { token } = w.session();
    w.advance(TOKEN_MAX_AGE_S - 1);
    expect(w.policy.resolve(token, 'read').state).toBe('active');
  });

  it('answers beyond_cap once the backstop passes, even with an active lock', async () => {
    const w = await createWorld();
    const { token, session } = w.session();
    w.advance(TOKEN_MAX_AGE_S - 10);
    w.acquire(session.id);
    w.advance(10);
    const err = await refusal(() => w.policy.resolve(token, 'read'));
    expect([err.status, err.details.reason]).toEqual([401, 'beyond_cap']);
  });

  it('answers revoked after revocation, and releases the lock', async () => {
    const w = await createWorld();
    const { token, session } = w.session();
    w.acquire(session.id);
    expect(w.policy.revokeSession(session.id)).toBe(true);
    const err = await refusal(() => w.policy.resolve(token, 'read'));
    expect([err.status, err.details.reason]).toEqual([401, 'revoked']);
    expect(w.repo.getLock(w.file.id)).toBeUndefined();
    expect(w.published.map((e) => e.type)).toContain('lock.released');
  });

  it('reports state "expired" while the session holds a lapsed lock', async () => {
    const w = await createWorld();
    const { token, session } = w.session();
    w.acquire(session.id);
    w.advance(LOCK_MAX_TIMEOUT_S + 1);
    expect(w.policy.resolve(token, 'lock').state).toBe('expired');
  });

  it('revokes every session of a user on logout', async () => {
    const w = await createWorld();
    const a1 = w.session('alice');
    const a2 = w.session('alice');
    const b = w.session('bob');
    w.acquire(a1.session.id);
    expect(w.policy.revokeUserSessions('alice')).toBe(2);
    expect((await refusal(() => w.policy.resolve(a2.token, 'read'))).details.reason).toBe('revoked');
    expect(w.policy.resolve(b.token, 'read').state).toBe('active');
    expect(w.repo.getLock(w.file.id)).toBeUndefined();
  });
});

describe('lock acquire', () => {
  it('records the base version, caps the timeout and refreshes last_seen_at', async () => {
    const w = await createWorld();
    const { session } = w.session();
    w.advance(5);
    const res = w.acquire(session.id, w.file.id, 99_999);
    expect(res.timeoutS).toBe(LOCK_MAX_TIMEOUT_S);
    expect(res.revived).toBe(false);
    expect(res.lockToken).toMatch(/^opaquelocktoken:[0-9a-f-]{36}$/);
    const row = w.repo.getLock(w.file.id)!;
    expect(row.base_version_id).toBe(w.file.current_version_id);
    expect(row.expires_at).toBe(w.now() + LOCK_MAX_TIMEOUT_S * 1000);
    expect(row.owner_xml).toBe('<D:owner xmlns:D="DAV:">x</D:owner>');
    expect(w.repo.getSession(session.id)!.last_seen_at).toBe(w.now());
    expect(w.published.at(-1)).toMatchObject({ type: 'lock.acquired', userId: 'alice' });
  });

  it('refuses a second session with 423 and names the holder', async () => {
    const w = await createWorld();
    const alice = w.session('alice');
    const bob = w.session('bob');
    w.acquire(alice.session.id);
    const err = await refusal(() => w.acquire(bob.session.id));
    expect(err.status).toBe(423);
    expect(err.details.heldBy).toEqual({ userId: 'alice', displayName: 'Alice', since: new Date(w.now()).toISOString() });
  });

  it('lets the same user take over the lock from their own dead session', async () => {
    const w = await createWorld();
    const crashed = w.session('alice');
    const reopened = w.session('alice');
    const old = w.acquire(crashed.session.id);
    const res = w.acquire(reopened.session.id);
    expect(res.lockToken).not.toBe(old.lockToken);
    expect(w.repo.getLock(w.file.id)!.session_id).toBe(reopened.session.id);
    expect(w.published.at(-1)).toMatchObject({ type: 'lock.acquired', tookOverFrom: crashed.session.id });
    // If the old session was alive after all, its next save is kept aside, not lost.
    expect((await w.put(crashed.session.id, old.lockToken)).kind).toBe('conflict');
  });

  it('refuses a new lock from the holder itself while it is active', async () => {
    const w = await createWorld();
    const { session } = w.session();
    w.acquire(session.id);
    expect((await refusal(() => w.acquire(session.id))).status).toBe(423);
  });

  it('replaces an expired lock', async () => {
    const w = await createWorld();
    const alice = w.session('alice');
    const bob = w.session('bob');
    w.acquire(alice.session.id);
    w.advance(LOCK_MAX_TIMEOUT_S);
    const res = w.acquire(bob.session.id);
    expect(w.repo.getLock(w.file.id)!.lock_token).toBe(res.lockToken);
    expect(w.repo.getLock(w.file.id)!.session_id).toBe(bob.session.id);
  });

  it('refuses revoked sessions', async () => {
    const w = await createWorld();
    const { session } = w.session();
    w.policy.revokeSession(session.id);
    expect((await refusal(() => w.acquire(session.id))).details.reason).toBe('revoked');
  });

  it('refuses a session for another file', async () => {
    const w = await createWorld();
    const other = await w.files.createFile({ name: 'Other.docx', body: body('x'), userId: 'bob', source: 'web' });
    const { session } = w.session('alice', other.id);
    expect((await refusal(() => w.acquire(session.id, w.file.id))).status).toBe(403);
  });
});

describe('lock refresh and revival', () => {
  it('extends an active lock', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    w.advance(30);
    const res = w.policy.lock({ fileId: w.file.id, sessionId: session.id, timeoutS: 45, lockToken });
    expect(res).toMatchObject({ lockToken, timeoutS: 45, revived: false });
    expect(w.repo.getLock(w.file.id)!.expires_at).toBe(w.now() + 45_000);
    expect(w.repo.getSession(session.id)!.last_seen_at).toBe(w.now());
    expect(w.published.at(-1)!.type).toBe('lock.refreshed');
  });

  it('revives an expired lock when all four conditions hold', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    w.advance(LOCK_MAX_TIMEOUT_S * 10);
    const res = w.policy.lock({ fileId: w.file.id, sessionId: session.id, timeoutS: 30, lockToken });
    expect(res.revived).toBe(true);
    expect(w.repo.getLock(w.file.id)!.expires_at).toBe(w.now() + 30_000);
    expect(w.published.at(-1)!.type).toBe('lock.revived');
  });

  it('condition 1: a revoked session is never revived', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    w.advance(LOCK_MAX_TIMEOUT_S + 1);
    w.policy.revokeSession(session.id);
    const err = await refusal(() => w.policy.lock({ fileId: w.file.id, sessionId: session.id, timeoutS: 30, lockToken }));
    expect([err.status, err.details.reason]).toEqual([401, 'revoked']);
  });

  it('condition 1: a session beyond the cap is not revived', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    w.advance(TOKEN_MAX_AGE_S);
    const err = await refusal(() => w.policy.lock({ fileId: w.file.id, sessionId: session.id, timeoutS: 30, lockToken }));
    expect([err.status, err.details.reason]).toEqual([401, 'beyond_cap']);
  });

  it('condition 2: a token that does not match the row gives 412', async () => {
    const w = await createWorld();
    const { session } = w.session();
    w.acquire(session.id);
    w.advance(LOCK_MAX_TIMEOUT_S + 1);
    const err = await refusal(() =>
      w.policy.lock({ fileId: w.file.id, sessionId: session.id, timeoutS: 30, lockToken: 'opaquelocktoken:other' }),
    );
    expect(err.status).toBe(412);
  });

  it('condition 2: no row at all (force-unlocked) gives 412', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    w.policy.forceUnlock(w.file.id);
    const err = await refusal(() => w.policy.lock({ fileId: w.file.id, sessionId: session.id, timeoutS: 30, lockToken }));
    expect(err.status).toBe(412);
  });

  it('condition 3: another session holding an active lock gives 423', async () => {
    const w = await createWorld();
    const alice = w.session('alice');
    const bob = w.session('bob');
    const { lockToken } = w.acquire(alice.session.id);
    w.advance(LOCK_MAX_TIMEOUT_S + 1);
    w.acquire(bob.session.id);
    const err = await refusal(() => w.policy.lock({ fileId: w.file.id, sessionId: alice.session.id, timeoutS: 30, lockToken }));
    expect(err.status).toBe(423);
    expect(err.details.heldBy?.userId).toBe('bob');
  });

  it('condition 4: a save after the lock expired gives 412', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    w.advance(LOCK_MAX_TIMEOUT_S + 1);
    await w.files.uploadVersion({ fileId: w.file.id, body: body('web'), userId: 'bob' });
    const err = await refusal(() => w.policy.lock({ fileId: w.file.id, sessionId: session.id, timeoutS: 30, lockToken }));
    expect(err.status).toBe(412);
  });

  it('a token is bound to its session', async () => {
    const w = await createWorld();
    const alice = w.session('alice');
    const bob = w.session('bob');
    const { lockToken } = w.acquire(alice.session.id);
    const err = await refusal(() => w.policy.lock({ fileId: w.file.id, sessionId: bob.session.id, timeoutS: 30, lockToken }));
    expect(err.status).toBe(423);
  });
});

describe('unlock', () => {
  it('releases the caller’s lock', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    w.policy.unlock(lockToken, session.id);
    expect(w.repo.getLock(w.file.id)).toBeUndefined();
    expect(w.published.at(-1)).toMatchObject({ type: 'lock.released', reason: 'unlock' });
  });

  it('refuses another session’s lock and ignores unknown tokens', async () => {
    const w = await createWorld();
    const alice = w.session('alice');
    const bob = w.session('bob');
    const { lockToken } = w.acquire(alice.session.id);
    expect((await refusal(() => w.policy.unlock(lockToken, bob.session.id))).status).toBe(403);
    expect(() => w.policy.unlock('opaquelocktoken:gone', bob.session.id)).not.toThrow();
  });
});

describe('PUT', () => {
  // Office reports every refused PUT as a successful save (observed with Word
  // 16.0.19127), so a known session's save is never refused: it is stored as
  // current, as a conflict side-version, or in quarantine.
  it('stores a save without a lock token as a conflict version, not the current one', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const res = await w.put(session.id, undefined, 'no lock');
    expect(res.kind).toBe('conflict');
    expect(w.repo.getFile(w.file.id)!.current_version_id).toBe(w.file.current_version_id);
    expect(w.repo.getVersion(res.versionId)).toMatchObject({ source: 'conflict', created_by: 'alice', dav_session_id: session.id });
    expect(w.published.at(-1)).toMatchObject({ type: 'version.created', source: 'conflict' });
  });

  it('accepts a save without an If header when the session holds the lock', async () => {
    const w = await createWorld();
    const { session } = w.session();
    w.acquire(session.id);
    const res = await w.put(session.id, undefined, 'edited');
    expect(res.kind).toBe('current');
    expect(w.repo.getFile(w.file.id)!.current_version_id).toBe(res.versionId);
  });

  it('saves under an active lock: new version, base version and expiry move forward', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    w.advance(20);
    const res = await w.put(session.id, lockToken, 'edited');
    expect(res.revived).toBe(false);
    expect(res.size).toBe(6);
    const file = w.repo.getFile(w.file.id)!;
    expect(file.current_version_id).toBe(res.versionId);
    expect(file.updated_by).toBe('alice');
    const version = w.repo.getVersion(res.versionId)!;
    expect(version).toMatchObject({ version_no: 2, source: 'dav', dav_session_id: session.id, created_by: 'alice' });
    expect(version.sha256).toMatch(/^[0-9a-f]{64}$/);
    const lock = w.repo.getLock(w.file.id)!;
    expect(lock.base_version_id).toBe(res.versionId);
    expect(lock.expires_at).toBe(w.now() + LOCK_MAX_TIMEOUT_S * 1000);
    expect(w.repo.getSession(session.id)!.last_seen_at).toBe(w.now());
    expect(w.published.at(-1)).toMatchObject({ type: 'version.created', versionNo: 2, source: 'dav' });
  });

  it('records the sha256 of the stored content', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    const res = await w.put(session.id, lockToken, 'hello');
    expect(w.repo.getVersion(res.versionId)!.sha256).toBe('2cf24dba5fb0a30e26e83b2ac5b9e29e1b161e5c1fa7425e73043362938b9824');
  });

  it('handles two quick saves in a row', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    const first = await w.put(session.id, lockToken, 'one');
    const second = await w.put(session.id, lockToken, 'two');
    expect(w.repo.listVersions(w.file.id).map((v) => v.id)).toEqual([second.versionId, first.versionId, w.file.current_version_id]);
  });

  it('revives an expired lock when the file is unchanged', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    w.advance(LOCK_MAX_TIMEOUT_S * 100);
    const res = await w.put(session.id, lockToken);
    expect(res.revived).toBe(true);
    expect(w.repo.getLock(w.file.id)!.expires_at).toBeGreaterThan(w.now());
    expect(w.published.map((e) => e.type)).toEqual(['lock.acquired', 'lock.revived', 'version.created']);
  });

  it('keeps a save as a conflict when another session saved after the lock expired', async () => {
    const w = await createWorld();
    const alice = w.session('alice');
    const bob = w.session('bob');
    const aliceLock = w.acquire(alice.session.id);
    w.advance(LOCK_MAX_TIMEOUT_S + 1);
    const bobLock = w.acquire(bob.session.id);
    const bobSave = await w.put(bob.session.id, bobLock.lockToken, 'bob');
    w.policy.unlock(bobLock.lockToken, bob.session.id);
    const res = await w.put(alice.session.id, aliceLock.lockToken, 'alice late');
    expect(res.kind).toBe('conflict');
    expect(w.repo.getFile(w.file.id)!.current_version_id).toBe(bobSave.versionId);
  });

  it('keeps a save as a conflict while another session holds the lock, and leaves that lock alone', async () => {
    const w = await createWorld();
    const alice = w.session('alice');
    const bob = w.session('bob');
    const { lockToken } = w.acquire(alice.session.id);
    const res = await w.put(bob.session.id, 'opaquelocktoken:bogus');
    expect(res.kind).toBe('conflict');
    expect(w.repo.getLock(w.file.id)!.lock_token).toBe(lockToken);
  });

  it('keeps a save after a force unlock as a conflict', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    w.policy.forceUnlock(w.file.id);
    expect((await w.put(session.id, lockToken)).kind).toBe('conflict');
  });

  it('quarantines a revoked session’s save instead of refusing it', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    w.policy.revokeSession(session.id);
    const res = await w.put(session.id, lockToken, 'late edits');
    expect(res.kind).toBe('quarantined');
    expect(w.repo.listVersions(w.file.id)).toHaveLength(1);
    const [q] = w.policy.listQuarantine();
    expect(q).toMatchObject({ session_id: session.id, user_id: 'alice', file_id: w.file.id, file_name: 'Plan.docx', reason: 'revoked', size: 10 });
    expect(existsSync(join(w.dir, 'blobs', q!.blob_key))).toBe(true);
    expect(w.published.at(-1)).toMatchObject({ type: 'upload.quarantined', reason: 'revoked' });
  });

  it('quarantines a save from a session past its cap', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    w.advance(TOKEN_MAX_AGE_S);
    expect((await w.put(session.id, lockToken)).kind).toBe('quarantined');
    expect(w.policy.listQuarantine()[0]!.reason).toBe('beyond_cap');
  });

  it('still refuses an unknown session id', async () => {
    const w = await createWorld();
    expect((await refusal(() => w.put('no-such-session', undefined))).status).toBe(401);
  });

  it('refuses uploads over MAX_UPLOAD_MB and keeps no blob', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    const big = Buffer.alloc(MAX_UPLOAD_BYTES + 1);
    const err = await refusal(() =>
      w.files.saveFromDav({ fileId: w.file.id, sessionId: session.id, lockToken, contentLength: undefined, body: Readable.from([big]) }),
    );
    expect(err.status).toBe(413);
    expect(readdirSync(join(w.dir, 'blobs', w.file.id))).toHaveLength(1);
  });

  it('re-checks the lock after the upload: a lock lost mid-upload gives a conflict, not the current version', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    // The lock is force-removed while the body is still arriving.
    const upload = new Readable({
      read() {
        w.policy.forceUnlock(w.file.id);
        this.push(Buffer.from('late'));
        this.push(null);
      },
    });
    const res = await w.files.saveFromDav({ fileId: w.file.id, sessionId: session.id, lockToken, contentLength: 4, body: upload });
    expect(res.kind).toBe('conflict');
    expect(w.repo.getFile(w.file.id)!.current_version_id).toBe(w.file.current_version_id);
    expect(readdirSync(join(w.dir, 'blobs', w.file.id))).toHaveLength(2);
  });

  it('writes every version as a new immutable blob', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const { lockToken } = w.acquire(session.id);
    const res = await w.put(session.id, lockToken, 'v2');
    expect(existsSync(join(w.dir, 'blobs', w.file.id, w.file.current_version_id!))).toBe(true);
    expect(existsSync(join(w.dir, 'blobs', w.file.id, res.versionId))).toBe(true);
  });
});

describe('Save As', () => {
  // Word's Save As to the DAV folder: HEAD 404, then a PUT to the new name with
  // no LOCK and no If header (observed). Refusing it made Word re-point the
  // document at the new URL and lose every later save.
  it('creates an unlocked file owned by the user that the session can then lock and save', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const created = await w.files.saveAsFromDav({ sessionId: session.id, name: 'Plan v2.docx', contentLength: 2, body: body('v2') });
    expect(created).toMatchObject({ kind: 'current', created: true });
    const file = w.repo.getFile(created.fileId)!;
    expect(file).toMatchObject({ name: 'Plan v2.docx', updated_by: 'alice' });
    // Word LOCKs the new name right after the PUT; a server-held lock made that 423.
    expect(w.repo.getLock(created.fileId)).toBeUndefined();
    expect(w.policy.findSessionFile(session.id, 'plan V2.docx')).toBe(created.fileId);

    const { lockToken } = w.acquire(session.id, created.fileId);
    expect((await w.put(session.id, lockToken, 'second', created.fileId)).kind).toBe('current');
    // A PUT to the same name again (Save As over it) also goes to that file, under the lock.
    const again = await w.files.saveAsFromDav({ sessionId: session.id, name: 'Plan v2.docx', contentLength: 5, body: body('again') });
    expect(again).toMatchObject({ kind: 'current', created: false, fileId: created.fileId });
    expect(w.repo.listVersions(created.fileId)).toHaveLength(3);
  });

  it('quarantines a Save As to a name that is not an Office file', async () => {
    const w = await createWorld();
    const { session } = w.session();
    const res = await w.files.saveAsFromDav({ sessionId: session.id, name: 'Plan.pdf', contentLength: 1, body: body('x') });
    expect(res.kind).toBe('quarantined');
    expect(w.policy.listQuarantine()[0]).toMatchObject({ reason: 'unsupported_name', file_name: 'Plan.pdf' });
    expect(w.repo.countFiles()).toBe(1);
  });

  it('quarantines a Save As from a revoked session', async () => {
    const w = await createWorld();
    const { session } = w.session();
    w.policy.revokeSession(session.id);
    const res = await w.files.saveAsFromDav({ sessionId: session.id, name: 'Plan v2.docx', contentLength: 1, body: body('x') });
    expect(res.kind).toBe('quarantined');
    expect(w.repo.countFiles()).toBe(1);
  });
});

describe('quarantine by token', () => {
  it('keeps the body of a save whose token was revoked', async () => {
    const w = await createWorld();
    const { token, session } = w.session();
    w.policy.revokeSession(session.id);
    const { quarantineId } = await w.files.quarantineFromDav({ token, fileName: 'Plan.docx', body: body('late') });
    expect(w.policy.listQuarantine()).toEqual([expect.objectContaining({ id: quarantineId, session_id: session.id, reason: 'revoked', size: 4 })]);
  });

  it('refuses an unknown token and keeps no blob', async () => {
    const w = await createWorld();
    const err = await refusal(() => w.files.quarantineFromDav({ token: 'nope', fileName: 'Plan.docx', body: body('x') }));
    expect(err.status).toBe(401);
    expect(w.policy.listQuarantine()).toEqual([]);
    await new Promise((r) => setTimeout(r, 20));
    expect(existsSync(join(w.dir, 'blobs', 'quarantine', 'by-token'))
      ? readdirSync(join(w.dir, 'blobs', 'quarantine', 'by-token'))
      : []).toEqual([]);
  });
});

describe('web upload', () => {
  it('is refused with 409 while the file is locked', async () => {
    const w = await createWorld();
    const { session } = w.session();
    w.acquire(session.id);
    expect((await refusal(() => w.files.uploadVersion({ fileId: w.file.id, body: body('x'), userId: 'bob' }))).status).toBe(409);
  });

  it('is allowed once the lock has expired', async () => {
    const w = await createWorld();
    const { session } = w.session();
    w.acquire(session.id);
    w.advance(LOCK_MAX_TIMEOUT_S);
    const version = await w.files.uploadVersion({ fileId: w.file.id, body: body('x'), userId: 'bob' });
    expect(version).toMatchObject({ version_no: 2, source: 'web' });
  });

  it('rejects non-Office files and unsafe names', async () => {
    const w = await createWorld();
    expect((await refusal(() => w.files.createFile({ name: 'a.pdf', body: body('x'), userId: 'bob', source: 'web' }))).status).toBe(400);
    expect((await refusal(() => w.files.createFile({ name: 'a<b>.docx', body: body('x'), userId: 'bob', source: 'web' }))).status).toBe(400);
    const file = await w.files.createFile({ name: 'C:\\fakepath\\דוח (1).docx', body: body('x'), userId: 'bob', source: 'web' });
    expect(file.name).toBe('דוח (1).docx');
  });
});
