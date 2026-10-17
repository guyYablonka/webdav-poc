import { createHash, randomBytes, randomUUID } from 'node:crypto';
import type { Intent, LockRequest, LockResponse, ResolveResponse } from '@office-dav/gateway-contract';
import type { EventBus } from '../events/bus.js';
import type { StoredBlob } from '../storage/blob-store.js';
import type { LockRow, Repo, SessionRow, VersionRow } from '../storage/repo.js';
import { PolicyError, notFound, sessionInvalid } from './errors.js';
import type { Permissions } from './permissions.js';

/** Milliseconds since the epoch. Injected so tests can move time. */
export type Clock = () => number;

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

  /** Loads a session for a gateway call: it must exist, be live, and (if given) belong to `fileId`. */
  requireSession(sessionId: string, fileId?: string): SessionRow {
    const session = this.repo.getSession(sessionId);
    if (!session) throw sessionInvalid('unknown');
    this.assertLive(session);
    if (fileId !== undefined && session.file_id !== fileId) {
      throw new PolicyError(403, 'forbidden', 'the DAV session is for another file');
    }
    return session;
  }

  /** Revokes a session and releases any lock it holds. Returns false if it was already revoked. */
  revokeSession(sessionId: string): boolean {
    const now = this.clock();
    const outcome = this.repo.transaction(() => {
      const session = this.repo.getSession(sessionId);
      if (!session || session.revoked_at !== null) return null;
      this.repo.revokeSession(sessionId, now);
      const lock = this.repo.getLockBySession(sessionId);
      if (lock) this.repo.deleteLock(lock.file_id);
      return { session, lock };
    });
    if (!outcome) return false;
    const { session, lock } = outcome;
    this.events.publish({ type: 'session.revoked', fileId: session.file_id, userId: session.user_id, sessionId, at: iso(now) });
    if (lock) {
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
    const lock = this.repo.transaction(() => {
      const session = this.requireSession(req.sessionId, req.fileId);
      this.requireWriter(session);
      const file = this.repo.getFile(req.fileId);
      if (!file?.current_version_id) throw notFound('file');
      const now = this.clock();
      const existing = this.repo.getLock(req.fileId);
      if (existing && existing.expires_at > now) throw this.lockedBy(existing);
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

  /** Admin action. The holder's next save fails with 412 because its token no longer matches. */
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

  /** Fast check before a DAV upload is streamed into storage. */
  checkDavWrite(fileId: string, sessionId: string, lockToken: string | undefined): void {
    const session = this.requireSession(sessionId, fileId);
    this.requireWriter(session);
    const now = this.clock();
    if (!lockToken) {
      const lock = this.repo.getLock(fileId);
      if (lock && lock.expires_at > now) throw this.lockedBy(lock);
      throw new PolicyError(423, 'locked', 'saving requires a lock');
    }
    this.checkLockForWrite(fileId, session.id, lockToken, now);
  }

  /**
   * Records an uploaded blob as the file's new version. The lock is checked
   * again in the same transaction, because it may have changed during the upload.
   */
  commitDavWrite(input: {
    fileId: string;
    sessionId: string;
    lockToken: string;
    versionId: string;
    blob: StoredBlob;
  }): { version: VersionRow; revived: boolean } {
    const now = this.clock();
    const result = this.repo.transaction(() => {
      const session = this.requireSession(input.sessionId, input.fileId);
      this.requireWriter(session);
      const { lock, revived } = this.checkLockForWrite(input.fileId, session.id, input.lockToken, now);
      const version = this.repo.addVersion({
        id: input.versionId,
        fileId: input.fileId,
        blobKey: input.blob.key,
        size: input.blob.size,
        sha256: input.blob.sha256,
        createdBy: session.user_id,
        createdAt: now,
        source: 'dav',
        davSessionId: session.id,
      });
      const expiresAt = now + lock.timeout_s * 1000;
      this.repo.setLockBaseVersion(input.fileId, version.id);
      this.repo.extendLock(input.fileId, lock.timeout_s, expiresAt);
      this.repo.touchSession(session.id, now);
      return { version, revived, lock, expiresAt };
    });
    if (result.revived) {
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
      source: 'dav',
      at: iso(now),
    });
    return { version: result.version, revived: result.revived };
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
