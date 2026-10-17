import Database from 'better-sqlite3';

export type Db = Database.Database;

/**
 * SQLite stand-in for the production PostgreSQL schema. Table and column names
 * match production. Timestamps are epoch milliseconds here; in PostgreSQL they
 * are timestamptz.
 */
const SCHEMA = `
CREATE TABLE IF NOT EXISTS users (
  id            TEXT PRIMARY KEY,
  display_name  TEXT NOT NULL,
  created_at    INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS files (
  id                  TEXT PRIMARY KEY,
  name                TEXT NOT NULL,
  mime                TEXT NOT NULL,
  current_version_id  TEXT,
  size                INTEGER NOT NULL DEFAULT 0,
  created_at          INTEGER NOT NULL,
  updated_at          INTEGER NOT NULL,
  updated_by          TEXT REFERENCES users(id)
);

CREATE TABLE IF NOT EXISTS dav_sessions (
  id              TEXT PRIMARY KEY,
  token_hash      TEXT NOT NULL UNIQUE,
  user_id         TEXT NOT NULL REFERENCES users(id),
  file_id         TEXT NOT NULL REFERENCES files(id),
  can_write       INTEGER NOT NULL CHECK (can_write IN (0, 1)),
  created_at      INTEGER NOT NULL,
  last_seen_at    INTEGER NOT NULL,
  max_expires_at  INTEGER NOT NULL,
  revoked_at      INTEGER,
  user_agent      TEXT
);
CREATE INDEX IF NOT EXISTS dav_sessions_user_idx ON dav_sessions(user_id);
CREATE INDEX IF NOT EXISTS dav_sessions_file_idx ON dav_sessions(file_id);

CREATE TABLE IF NOT EXISTS file_versions (
  id              TEXT PRIMARY KEY,
  file_id         TEXT NOT NULL REFERENCES files(id),
  version_no      INTEGER NOT NULL,
  blob_key        TEXT NOT NULL,
  size            INTEGER NOT NULL,
  sha256          TEXT NOT NULL,
  created_by      TEXT NOT NULL REFERENCES users(id),
  created_at      INTEGER NOT NULL,
  source          TEXT NOT NULL CHECK (source IN ('seed', 'web', 'dav')),
  dav_session_id  TEXT REFERENCES dav_sessions(id),
  UNIQUE (file_id, version_no)
);

CREATE TABLE IF NOT EXISTS file_locks (
  file_id          TEXT PRIMARY KEY REFERENCES files(id),
  lock_token       TEXT NOT NULL UNIQUE,
  session_id       TEXT NOT NULL REFERENCES dav_sessions(id),
  user_id          TEXT NOT NULL REFERENCES users(id),
  base_version_id  TEXT NOT NULL REFERENCES file_versions(id),
  owner_xml        TEXT,
  timeout_s        INTEGER NOT NULL,
  expires_at       INTEGER NOT NULL,
  created_at       INTEGER NOT NULL
);
`;

export function openDatabase(path: string): Db {
  const db = new Database(path);
  db.pragma('journal_mode = WAL');
  db.pragma('foreign_keys = ON');
  db.pragma('busy_timeout = 5000');
  db.exec(SCHEMA);
  return db;
}
