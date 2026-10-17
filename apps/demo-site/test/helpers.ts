import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Readable } from 'node:stream';
import { DavPolicy } from '../src/domain/dav-policy.js';
import { FileService } from '../src/domain/files.js';
import { everyoneCanWrite } from '../src/domain/permissions.js';
import { EventBus, type DemoEvent } from '../src/events/bus.js';
import { LocalDiskBlobStore } from '../src/storage/blob-store.js';
import { openDatabase } from '../src/storage/db.js';
import { Repo } from '../src/storage/repo.js';

export const LOCK_MAX_TIMEOUT_S = 60;
export const TOKEN_MAX_AGE_S = 24 * 3600;
export const MAX_UPLOAD_BYTES = 64 * 1024;

export function body(content: string | Buffer): Readable {
  return Readable.from([Buffer.from(content)]);
}

/** A policy over an in-memory database, a temp blob store and a clock the test controls. */
export async function createWorld() {
  let now = Date.parse('2026-03-01T09:00:00.000Z');
  const clock = () => now;
  const dir = mkdtempSync(join(tmpdir(), 'demo-site-test-'));
  const repo = new Repo(openDatabase(':memory:'));
  const events = new EventBus();
  const published: DemoEvent[] = [];
  events.subscribe((e) => published.push(e));
  const policy = new DavPolicy(
    repo,
    { tokenMaxAgeS: TOKEN_MAX_AGE_S, lockMaxTimeoutS: LOCK_MAX_TIMEOUT_S },
    events,
    everyoneCanWrite,
    clock,
  );
  const blobs = new LocalDiskBlobStore(join(dir, 'blobs'));
  const files = new FileService(repo, blobs, policy, events, { maxUploadBytes: MAX_UPLOAD_BYTES }, clock);
  repo.upsertUser('alice', 'Alice', now);
  repo.upsertUser('bob', 'Bob', now);
  const file = await files.createFile({ name: 'Plan.docx', body: body('v1'), userId: 'alice', source: 'seed' });
  published.length = 0;

  return {
    repo,
    policy,
    files,
    blobs,
    dir,
    file,
    published,
    advance(seconds: number) {
      now += seconds * 1000;
    },
    now: () => now,
    session(userId = 'alice', fileId = file.id) {
      return policy.createSession({ userId, fileId, userAgent: 'test' });
    },
    acquire(sessionId: string, fileId = file.id, timeoutS = LOCK_MAX_TIMEOUT_S) {
      return policy.lock({ fileId, sessionId, timeoutS, ownerXml: '<D:owner xmlns:D="DAV:">x</D:owner>' });
    },
    put(sessionId: string, lockToken: string | undefined, content = 'edit', fileId = file.id) {
      return files.saveFromDav({ fileId, sessionId, lockToken, contentLength: Buffer.byteLength(content), body: body(content) });
    },
  };
}
