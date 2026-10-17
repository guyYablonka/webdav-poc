import type { Db } from './db.js';

export interface UserRow {
  id: string;
  display_name: string;
  created_at: number;
}

export interface FileRow {
  id: string;
  name: string;
  mime: string;
  current_version_id: string | null;
  size: number;
  created_at: number;
  updated_at: number;
  updated_by: string | null;
}

export type VersionSource = 'seed' | 'web' | 'dav' | 'conflict';

export interface QuarantineRow {
  id: string;
  session_id: string;
  user_id: string;
  file_id: string | null;
  file_name: string;
  reason: string;
  blob_key: string;
  size: number;
  sha256: string;
  created_at: number;
}

export interface VersionRow {
  id: string;
  file_id: string;
  version_no: number;
  blob_key: string;
  size: number;
  sha256: string;
  created_by: string;
  created_at: number;
  source: VersionSource;
  dav_session_id: string | null;
}

export interface SessionRow {
  id: string;
  token_hash: string;
  user_id: string;
  file_id: string;
  can_write: 0 | 1;
  created_at: number;
  last_seen_at: number;
  max_expires_at: number;
  revoked_at: number | null;
  user_agent: string | null;
}

export interface LockRow {
  file_id: string;
  lock_token: string;
  session_id: string;
  user_id: string;
  base_version_id: string;
  owner_xml: string | null;
  timeout_s: number;
  expires_at: number;
  created_at: number;
}

export interface NewVersion {
  id: string;
  fileId: string;
  blobKey: string;
  size: number;
  sha256: string;
  createdBy: string;
  createdAt: number;
  source: VersionSource;
  davSessionId: string | null;
}

/**
 * Thin data access over the schema in db.ts. No policy lives here; callers
 * compose these calls inside `transaction` to keep each change atomic.
 */
export class Repo {
  constructor(readonly db: Db) {}

  /** Runs `fn` in a write transaction (BEGIN IMMEDIATE), so check-then-write steps are atomic. */
  transaction<T>(fn: () => T): T {
    return this.db.transaction(fn).immediate();
  }

  // users ------------------------------------------------------------------

  upsertUser(id: string, displayName: string, now: number): void {
    this.db
      .prepare(
        `INSERT INTO users (id, display_name, created_at) VALUES (?, ?, ?)
         ON CONFLICT (id) DO UPDATE SET display_name = excluded.display_name`,
      )
      .run(id, displayName, now);
  }

  getUser(id: string): UserRow | undefined {
    return this.db.prepare('SELECT * FROM users WHERE id = ?').get(id) as UserRow | undefined;
  }

  // files ------------------------------------------------------------------

  insertFile(file: Omit<FileRow, 'current_version_id' | 'size' | 'updated_by'>): void {
    this.db
      .prepare('INSERT INTO files (id, name, mime, size, created_at, updated_at) VALUES (?, ?, ?, 0, ?, ?)')
      .run(file.id, file.name, file.mime, file.created_at, file.updated_at);
  }

  getFile(id: string): FileRow | undefined {
    return this.db.prepare('SELECT * FROM files WHERE id = ?').get(id) as FileRow | undefined;
  }

  listFiles(): FileRow[] {
    return this.db.prepare('SELECT * FROM files ORDER BY name COLLATE NOCASE').all() as FileRow[];
  }

  countFiles(): number {
    return (this.db.prepare('SELECT COUNT(*) AS n FROM files').get() as { n: number }).n;
  }

  // versions ---------------------------------------------------------------

  /** Inserts the next version and makes it the file's current version. */
  addVersion(v: NewVersion): VersionRow {
    this.insertVersionRow(v);
    this.db
      .prepare('UPDATE files SET current_version_id = ?, size = ?, updated_at = ?, updated_by = ? WHERE id = ?')
      .run(v.id, v.size, v.createdAt, v.createdBy, v.fileId);
    return this.getVersion(v.id)!;
  }

  /** Inserts a version that does not become current (a conflict save). */
  addSideVersion(v: NewVersion): VersionRow {
    this.insertVersionRow(v);
    return this.getVersion(v.id)!;
  }

  private insertVersionRow(v: NewVersion): void {
    const next = this.db
      .prepare('SELECT COALESCE(MAX(version_no), 0) + 1 AS n FROM file_versions WHERE file_id = ?')
      .get(v.fileId) as { n: number };
    this.db
      .prepare(
        `INSERT INTO file_versions (id, file_id, version_no, blob_key, size, sha256, created_by, created_at, source, dav_session_id)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      )
      .run(v.id, v.fileId, next.n, v.blobKey, v.size, v.sha256, v.createdBy, v.createdAt, v.source, v.davSessionId);
  }

  getVersion(id: string): VersionRow | undefined {
    return this.db.prepare('SELECT * FROM file_versions WHERE id = ?').get(id) as VersionRow | undefined;
  }

  listVersions(fileId: string): VersionRow[] {
    return this.db
      .prepare('SELECT * FROM file_versions WHERE file_id = ? ORDER BY version_no DESC')
      .all(fileId) as VersionRow[];
  }

  countVersions(fileId: string): number {
    return (this.db.prepare('SELECT COUNT(*) AS n FROM file_versions WHERE file_id = ?').get(fileId) as { n: number }).n;
  }

  // sessions ---------------------------------------------------------------

  insertSession(s: SessionRow): void {
    this.db
      .prepare(
        `INSERT INTO dav_sessions (id, token_hash, user_id, file_id, can_write, created_at, last_seen_at, max_expires_at, revoked_at, user_agent)
         VALUES (@id, @token_hash, @user_id, @file_id, @can_write, @created_at, @last_seen_at, @max_expires_at, @revoked_at, @user_agent)`,
      )
      .run(s);
  }

  getSession(id: string): SessionRow | undefined {
    return this.db.prepare('SELECT * FROM dav_sessions WHERE id = ?').get(id) as SessionRow | undefined;
  }

  getSessionByTokenHash(tokenHash: string): SessionRow | undefined {
    return this.db.prepare('SELECT * FROM dav_sessions WHERE token_hash = ?').get(tokenHash) as SessionRow | undefined;
  }

  listSessionsForFile(fileId: string): SessionRow[] {
    return this.db
      .prepare('SELECT * FROM dav_sessions WHERE file_id = ? ORDER BY created_at DESC')
      .all(fileId) as SessionRow[];
  }

  listLiveSessionsForUser(userId: string): SessionRow[] {
    return this.db
      .prepare('SELECT * FROM dav_sessions WHERE user_id = ? AND revoked_at IS NULL')
      .all(userId) as SessionRow[];
  }

  touchSession(id: string, now: number): void {
    this.db.prepare('UPDATE dav_sessions SET last_seen_at = ? WHERE id = ?').run(now, id);
  }

  revokeSession(id: string, now: number): void {
    this.db.prepare('UPDATE dav_sessions SET revoked_at = ? WHERE id = ? AND revoked_at IS NULL').run(now, id);
  }

  // locks ------------------------------------------------------------------

  getLock(fileId: string): LockRow | undefined {
    return this.db.prepare('SELECT * FROM file_locks WHERE file_id = ?').get(fileId) as LockRow | undefined;
  }

  getLockByToken(lockToken: string): LockRow | undefined {
    return this.db.prepare('SELECT * FROM file_locks WHERE lock_token = ?').get(lockToken) as LockRow | undefined;
  }

  /** A session can hold several locks: its file and any files it created with Save As. */
  listLocksBySession(sessionId: string): LockRow[] {
    return this.db.prepare('SELECT * FROM file_locks WHERE session_id = ?').all(sessionId) as LockRow[];
  }

  // session files (Save As) ------------------------------------------------

  addSessionFile(sessionId: string, fileId: string, name: string, now: number): void {
    this.db
      .prepare('INSERT INTO dav_session_files (session_id, file_id, name, created_at) VALUES (?, ?, ?, ?)')
      .run(sessionId, fileId, name, now);
  }

  sessionOwnsFile(sessionId: string, fileId: string): boolean {
    return (
      this.db.prepare('SELECT 1 FROM dav_session_files WHERE session_id = ? AND file_id = ?').get(sessionId, fileId) !==
      undefined
    );
  }

  listSessionFiles(sessionId: string): Array<{ file_id: string; name: string }> {
    return this.db
      .prepare('SELECT file_id, name FROM dav_session_files WHERE session_id = ? ORDER BY created_at')
      .all(sessionId) as Array<{ file_id: string; name: string }>;
  }

  // quarantine -------------------------------------------------------------

  insertQuarantine(q: QuarantineRow): void {
    this.db
      .prepare(
        `INSERT INTO dav_quarantine (id, session_id, user_id, file_id, file_name, reason, blob_key, size, sha256, created_at)
         VALUES (@id, @session_id, @user_id, @file_id, @file_name, @reason, @blob_key, @size, @sha256, @created_at)`,
      )
      .run(q);
  }

  listQuarantine(): QuarantineRow[] {
    return this.db.prepare('SELECT * FROM dav_quarantine ORDER BY created_at DESC').all() as QuarantineRow[];
  }

  insertLock(lock: LockRow): void {
    this.db
      .prepare(
        `INSERT INTO file_locks (file_id, lock_token, session_id, user_id, base_version_id, owner_xml, timeout_s, expires_at, created_at)
         VALUES (@file_id, @lock_token, @session_id, @user_id, @base_version_id, @owner_xml, @timeout_s, @expires_at, @created_at)`,
      )
      .run(lock);
  }

  deleteLock(fileId: string): void {
    this.db.prepare('DELETE FROM file_locks WHERE file_id = ?').run(fileId);
  }

  extendLock(fileId: string, timeoutS: number, expiresAt: number): void {
    this.db.prepare('UPDATE file_locks SET timeout_s = ?, expires_at = ? WHERE file_id = ?').run(timeoutS, expiresAt, fileId);
  }

  setLockBaseVersion(fileId: string, versionId: string): void {
    this.db.prepare('UPDATE file_locks SET base_version_id = ? WHERE file_id = ?').run(versionId, fileId);
  }
}
