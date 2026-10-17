import { randomUUID } from 'node:crypto';
import type { Readable } from 'node:stream';
import type { FileInfo, PutContentResponse } from '@office-dav/gateway-contract';
import type { EventBus } from '../events/bus.js';
import { BlobTooLargeError, type BlobStore, type StoredBlob } from '../storage/blob-store.js';
import type { FileRow, Repo, SessionRow, VersionRow, VersionSource } from '../storage/repo.js';
import type { Clock, DavPolicy } from './dav-policy.js';
import { PolicyError, notFound } from './errors.js';
import { officeTypeFor, type OfficeType } from './office-types.js';

const iso = (ms: number) => new Date(ms).toISOString();

export interface UserRef {
  id: string;
  displayName: string;
}

export interface FileSummary {
  id: string;
  name: string;
  mime: string;
  type: OfficeType | null;
  size: number;
  versionCount: number;
  currentVersionNo: number;
  updatedAt: string;
  updatedBy: UserRef | null;
  lock: { user: UserRef; since: string; expiresAt: string; sessionId: string } | null;
}

export interface VersionView {
  id: string;
  versionNo: number;
  createdAt: string;
  createdBy: UserRef;
  size: number;
  sha256: string;
  source: VersionSource;
  isCurrent: boolean;
}

export interface SessionView {
  id: string;
  user: UserRef;
  createdAt: string;
  lastSeenAt: string;
  maxExpiresAt: string;
  revokedAt: string | null;
  userAgent: string | null;
  holdsLock: boolean;
}

export interface FileDetail extends FileSummary {
  versions: VersionView[];
  sessions: SessionView[];
}

export function payloadTooLarge(): PolicyError {
  return new PolicyError(413, 'payload_too_large', 'the upload exceeds MAX_UPLOAD_MB');
}

/** Files and their immutable versions. Lock and session rules live in DavPolicy. */
export class FileService {
  constructor(
    private readonly repo: Repo,
    private readonly blobs: BlobStore,
    private readonly policy: DavPolicy,
    private readonly events: EventBus,
    private readonly opts: { maxUploadBytes: number },
    private readonly clock: Clock = Date.now,
  ) {}

  requireFile(fileId: string): FileRow {
    const file = this.repo.getFile(fileId);
    if (!file?.current_version_id) throw notFound('file');
    return file;
  }

  /** The gateway's view of a file. */
  fileInfo(fileId: string): FileInfo {
    const file = this.requireFile(fileId);
    const lock = this.policy.activeLock(fileId);
    return {
      fileId: file.id,
      name: file.name,
      mime: file.mime,
      size: file.size,
      versionId: file.current_version_id!,
      updatedAt: iso(file.updated_at),
      createdAt: iso(file.created_at),
      lock: lock
        ? {
            lockToken: lock.lock_token,
            ownerXml: lock.owner_xml,
            expiresAt: iso(lock.expires_at),
            timeoutS: lock.timeout_s,
            sessionId: lock.session_id,
          }
        : null,
    };
  }

  async openVersion(fileId: string, versionId?: string): Promise<{ file: FileRow; version: VersionRow; stream: Readable }> {
    const file = this.requireFile(fileId);
    const version = this.repo.getVersion(versionId ?? file.current_version_id!);
    if (!version || version.file_id !== fileId) throw notFound('version');
    return { file, version, stream: await this.blobs.get(version.blob_key) };
  }

  /** New file from a web upload or the seed. */
  async createFile(input: { name: string; body: Readable; userId: string; source: 'seed' | 'web' }): Promise<FileRow> {
    const name = cleanFileName(input.name);
    const type = officeTypeFor(name);
    if (!type) throw new PolicyError(400, 'bad_request', 'only .docx, .xlsx and .pptx files are supported');

    const fileId = randomUUID();
    const versionId = randomUUID();
    const blob = await this.store(fileId, versionId, input.body);
    const now = this.clock();
    const version = this.withCleanup(blob, () =>
      this.repo.transaction(() => {
        this.repo.insertFile({ id: fileId, name, mime: type.mime, created_at: now, updated_at: now });
        return this.repo.addVersion({
          id: versionId,
          fileId,
          blobKey: blob.key,
          size: blob.size,
          sha256: blob.sha256,
          createdBy: input.userId,
          createdAt: now,
          source: input.source,
          davSessionId: null,
        });
      }),
    );
    this.events.publish({ type: 'file.created', fileId, name, userId: input.userId, at: iso(now) });
    this.publishVersion(version);
    return this.repo.getFile(fileId)!;
  }

  /** New version from a web upload. Refused with 409 while someone holds an active lock. */
  async uploadVersion(input: { fileId: string; body: Readable; userId: string }): Promise<VersionRow> {
    this.requireFile(input.fileId);
    this.policy.assertNotLocked(input.fileId);
    const versionId = randomUUID();
    const blob = await this.store(input.fileId, versionId, input.body);
    const version = this.withCleanup(blob, () =>
      this.repo.transaction(() => {
        this.policy.assertNotLocked(input.fileId);
        return this.repo.addVersion({
          id: versionId,
          fileId: input.fileId,
          blobKey: blob.key,
          size: blob.size,
          sha256: blob.sha256,
          createdBy: input.userId,
          createdAt: this.clock(),
          source: 'web',
          davSessionId: null,
        });
      }),
    );
    this.publishVersion(version);
    return version;
  }

  /** A save from Office through the gateway's PUT /files/:fileId/content. */
  async saveFromDav(input: {
    fileId: string;
    sessionId: string;
    lockToken: string | undefined;
    contentLength: number | undefined;
    body: Readable;
  }): Promise<PutContentResponse> {
    this.requireFile(input.fileId);
    // Refuse before reading the body when the outcome is already known.
    this.policy.checkDavWrite(input.fileId, input.sessionId, input.lockToken);
    if (input.contentLength !== undefined && input.contentLength > this.opts.maxUploadBytes) throw payloadTooLarge();

    const versionId = randomUUID();
    const blob = await this.store(input.fileId, versionId, input.body);
    const { version, revived } = this.withCleanup(blob, () =>
      this.policy.commitDavWrite({
        fileId: input.fileId,
        sessionId: input.sessionId,
        lockToken: input.lockToken!,
        versionId,
        blob,
      }),
    );
    return { versionId: version.id, size: version.size, updatedAt: iso(version.created_at), revived };
  }

  listSummaries(): FileSummary[] {
    return this.repo
      .listFiles()
      .filter((f) => f.current_version_id)
      .map((f) => this.summary(f));
  }

  detail(fileId: string): FileDetail {
    const file = this.requireFile(fileId);
    const summary = this.summary(file);
    const lock = this.policy.activeLock(fileId);
    return {
      ...summary,
      versions: this.repo.listVersions(fileId).map((v) => ({
        id: v.id,
        versionNo: v.version_no,
        createdAt: iso(v.created_at),
        createdBy: this.userRef(v.created_by),
        size: v.size,
        sha256: v.sha256,
        source: v.source,
        isCurrent: v.id === file.current_version_id,
      })),
      sessions: this.repo.listSessionsForFile(fileId).map((s) => this.sessionView(s, lock?.session_id)),
    };
  }

  private summary(file: FileRow): FileSummary {
    const lock = this.policy.activeLock(file.id);
    const current = this.repo.getVersion(file.current_version_id!);
    return {
      id: file.id,
      name: file.name,
      mime: file.mime,
      type: officeTypeFor(file.name),
      size: file.size,
      versionCount: this.repo.countVersions(file.id),
      currentVersionNo: current?.version_no ?? 0,
      updatedAt: iso(file.updated_at),
      updatedBy: file.updated_by ? this.userRef(file.updated_by) : null,
      lock: lock
        ? {
            user: this.userRef(lock.user_id),
            since: iso(lock.created_at),
            expiresAt: iso(lock.expires_at),
            sessionId: lock.session_id,
          }
        : null,
    };
  }

  private sessionView(s: SessionRow, lockSessionId: string | undefined): SessionView {
    return {
      id: s.id,
      user: this.userRef(s.user_id),
      createdAt: iso(s.created_at),
      lastSeenAt: iso(s.last_seen_at),
      maxExpiresAt: iso(s.max_expires_at),
      revokedAt: s.revoked_at === null ? null : iso(s.revoked_at),
      userAgent: s.user_agent,
      holdsLock: s.id === lockSessionId,
    };
  }

  private userRef(userId: string): UserRef {
    return { id: userId, displayName: this.repo.getUser(userId)?.display_name ?? userId };
  }

  private publishVersion(version: VersionRow): void {
    this.events.publish({
      type: 'version.created',
      fileId: version.file_id,
      versionId: version.id,
      versionNo: version.version_no,
      userId: version.created_by,
      source: version.source,
      at: iso(version.created_at),
    });
  }

  private async store(fileId: string, versionId: string, body: Readable): Promise<StoredBlob> {
    try {
      return await this.blobs.put(`${fileId}/${versionId}`, body, { maxBytes: this.opts.maxUploadBytes });
    } catch (err) {
      if (err instanceof BlobTooLargeError) throw payloadTooLarge();
      throw err;
    }
  }

  /** Runs the database step for a stored blob, deleting the blob if that step fails. */
  private withCleanup<T>(blob: StoredBlob, fn: () => T): T {
    try {
      return fn();
    } catch (err) {
      this.blobs.delete(blob.key).catch(() => {});
      throw err;
    }
  }

}

/** Keeps the base name only and rejects names Windows or Office cannot use. */
export function cleanFileName(raw: string): string {
  const base = raw.split(/[\\/]/).pop()!.normalize('NFC').trim();
  if (!base || base.length > 200 || /[\u0000-\u001f<>:"|?*]/.test(base) || base.startsWith('.')) {
    throw new PolicyError(400, 'bad_request', 'invalid file name');
  }
  return base;
}
