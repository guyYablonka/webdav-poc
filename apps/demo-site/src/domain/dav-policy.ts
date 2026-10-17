import { createHash, randomBytes, randomUUID } from 'node:crypto';
import type { Intent, LockRequest, LockResponse, ResolveResponse } from '@office-dav/gateway-contract';
import type { EventBus } from '../events/bus.js';
import type { StoredBlob } from '../storage/blob-store.js';
import type { LockRow, QuarantineRow, Repo, SessionRow, VersionRow } from '../storage/repo.js';
import { PolicyError, notFound, sessionInvalid } from './errors.js';
import type { Permissions } from './permissions.js';

/** Milliseconds since the epoch. Injected so tests can move time. */
export type Clock = () => number;

/** Where a DAV save goes. See SaveKind in the gateway contract. */
export type SavePlan =
  | { kind: 'current'; lock: LockRow; revived: boolean }
  | { kind: 'conflict'; reason: string }
  | { kind: 'quarantined'; reason: string };

export type SaveAsPlan =
  | { kind: 'existing'; fileId: string }
  | { kind: 'new' }
  | { kind: 'quarantined'; reason: string };

export type DavWriteResult =
  | { kind: 'current'; version: VersionRow; revived: boolean }
  | { kind: 'conflict'; version: VersionRow; reason: string; revived: false }
  | { kind: 'quarantined'; id: string; reason: string; revived: false };

export interface PolicyConfig {
  /** Backstop lifetime of a session token (TOKEN_MAX_AGE_S). */
  tokenMaxAgeS: number;
  /** Cap for a granted lock timeout (LOCK_MAX_TIMEOUT_S). */
  lockMaxTimeoutS: number;
}

export function hashToken(token: string): string {
  return createHash('sha256').update(token).digest('hex');
}

const iso = (ms: number) => new Date(ms).toISOString();

/**
 * The DAV session, lock and save rules. Every check-then-write runs in one
 * transaction, and events are published only after it commits.
 *
 * Sessions
 *   - Tokens are 32 random bytes (base64url); only sha256(token) is stored.
 *   - No idle timeout. `max_expires_at` (created + TOKEN_MAX_AGE_S) is a backstop.
 *   - Revoked always means 401 and cannot be revived. Beyond the cap is 401 beyond_cap.
 *
 * Locks
 *   - Acquire succeeds when there is no lock row or the row has expired; else 423.
 *   - Refresh, or a PUT, under an expired lock revives it when: the session is
 *     live (1), the token matches the row (2), no other session holds an
 *     active lock (3, else 423), and the file has not changed since the lock's
 *     base version (4, else 412).
 *   - Every successful LOCK or PUT refreshes last_seen_at and expires_at.
 */
export class DavPolicy {
  constructor(
    private readonly repo: Repo,
    private readonly config: PolicyConfig,
    private readonly events: EventBus,
    private readonly permissions: Permissions,
    private readonly clock: Clock = Date.now,
  ) {}

  // -------------------------------------------------------------------------
  // Sessions
  // -------------------------------------------------------------------------

  createSession(input: { userId: string; fileId: string; userAgent?: string | null }): {
    token: string;
    session: SessionRow;
  } {
    if (!this.repo.getFile(input.fileId)) throw notFound('file');
    if (!this.permissions.canWrite(input.userId, input.fileId)) {
      throw new PolicyError(403, 'forbidden', 'no write permission for this file');
    }
    const now = this.clock();
    const token = randomBytes(32).toString('base64url');
    const session: SessionRow = {
      id: randomUUID(),
      token_hash: hashToken(token),
      user_id: input.userId,
      file_id: input.fileId,
      can_write: 1,
      created_at: now,
      last_seen_at: now,
      max_expires_at: now + this.config.tokenMaxAgeS * 1000,
      revoked_at: null,
      user_agent: input.userAgent?.slice(0, 300) ?? null,
    };
    this.repo.insertSession(session);
    return { token, session };
  }

  /** Idempotent and read-only, so dav-service may cache the answer. */
  resolve(token: string, _intent: Intent): ResolveResponse {
    const session = this.repo.getSessionByTokenHash(hashToken(token));
    if (!session) throw sessionInvalid('unknown');
    this.assertLive(session);
    const lock = this.repo.getLock(session.file_id);
    const lapsed = lock !== undefined && lock.session_id === session.id && lock.expires_at <= this.clock();
    return {
      sessionId: session.id,
      userId: session.user_id,
      displayName: this.displayName(session.user_id),
      fileId: session.file_id,
      canWrite: this.isWriter(session),
      state: lapsed ? 'expired' : 'active',
    };
  }

  /**
   * Loads a session for a gateway call: it must exist, be live, and (if given)
   * cover `fileId`, meaning its own file or a file it created with Save As.
   */
  requireSession(sessionId: string, fileId?: string): SessionRow {
    const session = this.repo.getSession(sessionId);
    if (!session) throw sessionInvalid('unknown');
    this.assertLive(session);
    if (fileId !== undefined && !this.covers(session, fileId)) {
      throw new PolicyError(403, 'forbidden', 'the DAV session is for another file');
    }
    return session;
  }

  /** Loads a session row whatever its state (for attributing quarantined saves). */
  requireSessionRow(sessionId: string): SessionRow {
    const session = this.repo.getSession(sessionId);
    if (!session) throw sessionInvalid('unknown');
    return session;
  }

  /** Revokes a session and releases every lock it holds. Returns false if it was already revoked. */
  revokeSession(sessionId: string): boolean {
    const now = this.clock();
    const outcome = this.repo.transaction(() => {
      const session = this.repo.getSession(sessionId);
      if (!session || session.revoked_at !== null) return null;
      this.repo.revokeSession(sessionId, now);
      const locks = this.repo.listLocksBySession(sessionId);
      for (const lock of locks) this.repo.deleteLock(lock.file_id);
      return { session, locks };
    });
    if (!outcome) return false;
    const { session, locks } = outcome;
    this.events.publish({ type: 'session.revoked', fileId: session.file_id, userId: session.user_id, sessionId, at: iso(now) });
    for (const lock of locks) {
      this.events.publish({ type: 'lock.released', fileId: lock.file_id, userId: lock.user_id, reason: 'session_revoked', at: iso(now) });
    }
    return true;
  }

  /** Used on logout. */
  revokeUserSessions(userId: string): number {
    let revoked = 0;
    for (const session of this.repo.listLiveSessionsForUser(userId)) {
      if (this.revokeSession(session.id)) revoked++;
    }
    return revoked;
  }

  // -------------------------------------------------------------------------
  // Locks
  // -------------------------------------------------------------------------

  lock(req: LockRequest): LockResponse {
    return req.lockToken === undefined
      ? this.acquireLock(req)
      : this.refreshLock({ ...req, lockToken: req.lockToken });
  }

  acquireLock(req: Omit<LockRequest, 'lockToken'>): LockResponse {
    const timeoutS = this.grantTimeout(req.timeoutS);
    let tookOverFrom: string | undefined;
    const lock = this.repo.transaction(() => {
      const session = this.requireSession(req.sessionId, req.fileId);
      this.requireWriter(session);
      const file = this.repo.getFile(req.fileId);
      if (!file?.current_version_id) throw notFound('file');
      const now = this.clock();
      const existing = this.repo.getLock(req.fileId);
      if (existing && existing.expires_at > now) {
        // Same-user takeover: after Office crashes, the user's next "Edit" is a
        // new session while the dead one still holds the lock. Refusing meant
        // Word told the user the file was "locked by another user" (their own
        // lock) for up to the lock timeout. If the old session is in fact still
        // alive, its next save becomes a conflict version, so nothing is lost.
        if (existing.user_id !== session.user_id || existing.session_id === session.id) throw this.lockedBy(existing);
        tookOverFrom = existing.session_id;
      }
      if (existing) this.repo.deleteLock(req.fileId);
      const lock: LockRow = {
        file_id: req.fileId,
        lock_token: `opaquelocktoken:${randomUUID()}`,
        session_id: session.id,
        user_id: session.user_id,
        base_version_id: file.current_version_id,
        owner_xml: req.ownerXml ?? null,
        timeout_s: timeoutS,
        expires_at: now + timeoutS * 1000,
        created_at: now,
      };
      this.repo.insertLock(lock);
      this.repo.touchSession(session.id, now);
      return lock;
    });
    this.events.publish({
      type: 'lock.acquired',
      fileId: lock.file_id,
      userId: lock.user_id,
      sessionId: lock.session_id,
      expiresAt: iso(lock.expires_at),
      at: iso(lock.created_at),
      ...(tookOverFrom ? { tookOverFrom } : {}),
    });
    return { lockToken: lock.lock_token, expiresAt: iso(lock.expires_at), timeoutS, revived: false };
  }

  refreshLock(req: LockRequest & { lockToken: string }): LockResponse {
    const timeoutS = this.grantTimeout(req.timeoutS);
    const now = this.clock();
    const { lock, revived } = this.repo.transaction(() => {
      const session = this.requireSession(req.sessionId, req.fileId); // condition 1
      this.requireWriter(session);
      const checked = this.checkLockForWrite(req.fileId, session.id, req.lockToken, now); // conditions 2-4
      const expiresAt = now + timeoutS * 1000;
      this.repo.extendLock(req.fileId, timeoutS, expiresAt);
      this.repo.touchSession(session.id, now);
      return { lock: { ...checked.lock, timeout_s: timeoutS, expires_at: expiresAt }, revived: checked.revived };
    });
    this.events.publish({
      type: revived ? 'lock.revived' : 'lock.refreshed',
      fileId: lock.file_id,
      userId: lock.user_id,
      sessionId: lock.session_id,
      expiresAt: iso(lock.expires_at),
      at: iso(now),
    });
    return { lockToken: lock.lock_token, expiresAt: iso(lock.expires_at), timeoutS, revived };
  }

  /** Releases the caller's lock. Unknown tokens succeed: the lock is already gone. */
  unlock(lockToken: string, sessionId: string): void {
    const now = this.clock();
    const released = this.repo.transaction(() => {
      const session = this.requireSession(sessionId);
      const lock = this.repo.getLockByToken(lockToken);
      if (!lock) return null;
      if (lock.session_id !== session.id) throw new PolicyError(403, 'forbidden', 'the lock belongs to another session');
      this.repo.deleteLock(lock.file_id);
      return lock;
    });
    if (released) {
      this.events.publish({ type: 'lock.released', fileId: released.file_id, userId: released.user_id, reason: 'unlock', at: iso(now) });
    }
  }

  /** Admin action. The holder's next save becomes a conflict version because its token no longer matches. */
  forceUnlock(fileId: string): boolean {
    const now = this.clock();
    const released = this.repo.transaction(() => {
      const lock = this.repo.getLock(fileId);
      if (lock) this.repo.deleteLock(fileId);
      return lock;
    });
    if (!released) return false;
    this.events.publish({ type: 'lock.released', fileId, userId: released.user_id, reason: 'force_unlock', at: iso(now) });
    return true;
  }

  /** The current lock if it has not expired. */
  activeLock(fileId: string): LockRow | undefined {
    const lock = this.repo.getLock(fileId);
    return lock && lock.expires_at > this.clock() ? lock : undefined;
  }

  // -------------------------------------------------------------------------
  // Saves
  // -------------------------------------------------------------------------

  /**
   * Where a DAV save will go, decided before the body is read. Never refuses a
   * known session: Office treats any refusal as success and drops the edits
   * (see SaveKind in the contract). Throws only for an unknown session id.
   */
  planDavWrite(fileId: string, sessionId: string, lockToken: string | undefined): SavePlan {
    const session = this.repo.getSession(sessionId);
    if (!session) throw sessionInvalid('unknown');
    return this.decideWrite(session, fileId, lockToken, this.clock());
  }

  /**
   * Records an uploaded blob for a DAV save. The decision is made again in the
   * same transaction, because the lock or the session may have changed during
   * the upload: the result can be a current version, a conflict version, or a
   * quarantine entry, but the bytes are always kept.
   */
  commitDavWrite(input: {
    fileId: string;
    fileName: string;
    sessionId: string;
    lockToken: string | undefined;
    versionId: string;
    blob: StoredBlob;
    /** Quarantine regardless of the lock, with this reason (Save As the session may not make). */
    forceQuarantine?: string;
  }): DavWriteResult {
    const now = this.clock();
    const result = this.repo.transaction(() => {
      const session = this.repo.getSession(input.sessionId);
      if (!session) throw sessionInvalid('unknown');
      const plan: SavePlan = input.forceQuarantine
        ? { kind: 'quarantined', reason: input.forceQuarantine }
        : this.decideWrite(session, input.fileId, input.lockToken, now);
      if (plan.kind === 'quarantined') {
        this.insertQuarantine(session, input.fileId, input.fileName, plan.reason, input.versionId, input.blob, now);
        return { kind: plan.kind, reason: plan.reason, session };
      }
      const row = {
        id: input.versionId,
        fileId: input.fileId,
        blobKey: input.blob.key,
        size: input.blob.size,
        sha256: input.blob.sha256,
        createdBy: session.user_id,
        createdAt: now,
        davSessionId: session.id,
      };
      this.repo.touchSession(session.id, now);
      if (plan.kind === 'conflict') {
        const version = this.repo.addSideVersion({ ...row, source: 'conflict' });
        return { kind: plan.kind, reason: plan.reason, session, version };
      }
      const version = this.repo.addVersion({ ...row, source: 'dav' });
      const expiresAt = now + plan.lock.timeout_s * 1000;
      this.repo.setLockBaseVersion(input.fileId, version.id);
      this.repo.extendLock(input.fileId, plan.lock.timeout_s, expiresAt);
      return { kind: plan.kind, session, version, revived: plan.revived, lock: plan.lock, expiresAt };
    });

    if (result.kind === 'quarantined') {
      this.events.publish({
        type: 'upload.quarantined',
        fileId: input.fileId,
        userId: result.session.user_id,
        sessionId: result.session.id,
        reason: result.reason,
        at: iso(now),
      });
      return { kind: 'quarantined', id: input.versionId, reason: result.reason, revived: false };
    }
    if (result.kind === 'current' && result.revived) {
      this.events.publish({
        type: 'lock.revived',
        fileId: input.fileId,
        userId: result.lock.user_id,
        sessionId: input.sessionId,
        expiresAt: iso(result.expiresAt),
        at: iso(now),
      });
    }
    this.events.publish({
      type: 'version.created',
      fileId: input.fileId,
      versionId: result.version.id,
      versionNo: result.version.version_no,
      userId: result.version.created_by,
      source: result.version.source,
      at: iso(now),
    });
    return result.kind === 'current'
      ? { kind: 'current', version: result.version, revived: result.revived }
      : { kind: 'conflict', version: result.version, reason: result.reason, revived: false };
  }

  /**
   * Stores a save whose token no longer resolves (revoked or beyond its cap)
   * in quarantine. Unknown tokens are refused: there is no user to attribute
   * the bytes to, and accepting them would let anyone fill the store.
   */
  quarantineByToken(token: string, fileName: string, id: string, blob: StoredBlob): { sessionId: string; reason: string } {
    const now = this.clock();
    const session = this.repo.getSessionByTokenHash(hashToken(token));
    if (!session) throw sessionInvalid('unknown');
    const reason = session.revoked_at !== null ? 'revoked' : now >= session.max_expires_at ? 'beyond_cap' : 'session_invalid';
    this.repo.transaction(() => this.insertQuarantine(session, session.file_id, fileName, reason, id, blob, now));
    this.events.publish({
      type: 'upload.quarantined',
      fileId: session.file_id,
      userId: session.user_id,
      sessionId: session.id,
      reason,
      at: iso(now),
    });
    return { sessionId: session.id, reason };
  }

  /**
   * Save As: where a PUT to `name` (not the session's file) should go.
   *   existing     the session already created this name; save to it as usual
   *   new          create a file with this name in the session's collection
   *   quarantined  the session may not write, or the name is not an Office file
   */
  planSaveAs(sessionId: string, name: string, supported: boolean): SaveAsPlan {
    const session = this.repo.getSession(sessionId);
    if (!session) throw sessionInvalid('unknown');
    const refusal = this.writeRefusal(session, this.clock());
    if (refusal) return { kind: 'quarantined', reason: refusal };
    const existing = this.findSessionFile(sessionId, name);
    if (existing) return { kind: 'existing', fileId: existing };
    if (!supported) return { kind: 'quarantined', reason: 'unsupported_name' };
    return { kind: 'new' };
  }

  /**
   * Links a file created by Save As to the session, so the session's token
   * covers it. No lock is taken here: Word sends PUT (no If header), then LOCK
   * on the new name straight away (observed). A server-held lock made that
   * LOCK fail with 423, and Word then DELETEd the new file and lost the save.
   */
  registerSaveAs(sessionId: string, fileId: string, name: string): void {
    const now = this.clock();
    this.repo.transaction(() => {
      this.requireSession(sessionId);
      const file = this.repo.getFile(fileId);
      if (!file?.current_version_id) throw notFound('file');
      this.repo.addSessionFile(sessionId, fileId, name, now);
    });
  }

  /** The id of the file this session created under `name`, if any. Same name rules as the URL space. */
  findSessionFile(sessionId: string, name: string): string | undefined {
    const wanted = name.normalize('NFC').toLowerCase();
    return this.repo.listSessionFiles(sessionId).find((f) => f.name.normalize('NFC').toLowerCase() === wanted)?.file_id;
  }

  /** Admin view of quarantined saves. */
  listQuarantine(): QuarantineRow[] {
    return this.repo.listQuarantine();
  }

  /** Web uploads are refused with 409 while anyone holds an active lock. */
  assertNotLocked(fileId: string): void {
    const lock = this.activeLock(fileId);
    if (lock) {
      throw new PolicyError(409, 'conflict', `being edited by ${this.displayName(lock.user_id)}`, {
        heldBy: this.holder(lock),
      });
    }
  }

  // -------------------------------------------------------------------------

  /**
   * Is `lockToken` the caller's usable lock? Covers the plain case (an active
   * lock it holds) and revival conditions 2-4. Condition 1 is the caller's job.
   */
  private checkLockForWrite(
    fileId: string,
    sessionId: string,
    lockToken: string,
    now: number,
  ): { lock: LockRow; revived: boolean } {
    const lock = this.repo.getLock(fileId);
    const ours = lock !== undefined && lock.lock_token === lockToken && lock.session_id === sessionId;
    if (!ours) {
      // Condition 3: another session holds an active lock.
      if (lock && lock.expires_at > now) throw this.lockedBy(lock);
      // Condition 2: the token does not match the lock row, or there is none.
      throw new PolicyError(412, 'precondition_failed', 'the lock token does not match the current lock');
    }
    if (lock.expires_at > now) return { lock, revived: false };
    // Condition 4: nobody saved since this lock's base version.
    const file = this.repo.getFile(fileId);
    if (file?.current_version_id !== lock.base_version_id) {
      throw new PolicyError(412, 'precondition_failed', 'the file changed after the lock expired');
    }
    return { lock, revived: true };
  }

  /**
   * The save decision shared by planDavWrite and commitDavWrite.
   * A missing lock token is accepted when the session itself holds the file's
   * lock (Office sometimes saves under its lock without sending the If header).
   */
  private decideWrite(session: SessionRow, fileId: string, lockToken: string | undefined, now: number): SavePlan {
    const refusal = this.writeRefusal(session, now);
    if (refusal) return { kind: 'quarantined', reason: refusal };
    if (!this.covers(session, fileId)) return { kind: 'quarantined', reason: 'other_file' };
    const token = lockToken ?? this.repo.getLock(fileId)?.lock_token;
    if (!token) return { kind: 'conflict', reason: 'no_lock' };
    try {
      const { lock, revived } = this.checkLockForWrite(fileId, session.id, token, now);
      return { kind: 'current', lock, revived };
    } catch (err) {
      if (err instanceof PolicyError) return { kind: 'conflict', reason: err.status === 423 ? 'locked_by_other' : 'lock_lost' };
      throw err;
    }
  }

  /** Why this session may not write at all, or null when it may. */
  private writeRefusal(session: SessionRow, now: number): string | null {
    if (session.revoked_at !== null) return 'revoked';
    if (now >= session.max_expires_at) return 'beyond_cap';
    if (!this.isWriter(session)) return 'forbidden';
    return null;
  }

  private covers(session: SessionRow, fileId: string): boolean {
    return session.file_id === fileId || this.repo.sessionOwnsFile(session.id, fileId);
  }

  private insertQuarantine(
    session: SessionRow,
    fileId: string | null,
    fileName: string,
    reason: string,
    id: string,
    blob: StoredBlob,
    now: number,
  ): void {
    this.repo.insertQuarantine({
      id,
      session_id: session.id,
      user_id: session.user_id,
      file_id: fileId && this.repo.getFile(fileId) ? fileId : null,
      file_name: fileName,
      reason,
      blob_key: blob.key,
      size: blob.size,
      sha256: blob.sha256,
      created_at: now,
    });
  }

  private assertLive(session: SessionRow): void {
    if (session.revoked_at !== null) throw sessionInvalid('revoked');
    if (this.clock() >= session.max_expires_at) throw sessionInvalid('beyond_cap');
  }

  private isWriter(session: SessionRow): boolean {
    return session.can_write === 1 && this.permissions.canWrite(session.user_id, session.file_id);
  }

  private requireWriter(session: SessionRow): void {
    if (!this.isWriter(session)) throw new PolicyError(403, 'forbidden', 'the DAV session is read-only');
  }

  private grantTimeout(requestedS: number): number {
    return Math.max(1, Math.min(Math.floor(requestedS), this.config.lockMaxTimeoutS));
  }

  private displayName(userId: string): string {
    return this.repo.getUser(userId)?.display_name ?? userId;
  }

  private holder(lock: LockRow) {
    return { userId: lock.user_id, displayName: this.displayName(lock.user_id), since: iso(lock.created_at) };
  }

  private lockedBy(lock: LockRow): PolicyError {
    return new PolicyError(423, 'locked', `locked by ${this.displayName(lock.user_id)}`, { heldBy: this.holder(lock) });
  }
}
