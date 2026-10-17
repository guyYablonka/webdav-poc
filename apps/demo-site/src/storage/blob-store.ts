import { createHash, randomUUID } from 'node:crypto';
import { createReadStream, createWriteStream } from 'node:fs';
import { access, mkdir, rename, rm } from 'node:fs/promises';
import { dirname, join, resolve, sep } from 'node:path';
import { Transform, type Readable, type TransformCallback } from 'node:stream';
import { pipeline } from 'node:stream/promises';

export interface StoredBlob {
  key: string;
  size: number;
  /** Hex SHA-256 of the content, computed while streaming. */
  sha256: string;
}

/**
 * Immutable blob storage. Every version is written once under a fresh key and
 * never modified. S3 replaces the local-disk implementation in production.
 */
export interface BlobStore {
  put(key: string, body: Readable, opts?: { maxBytes?: number }): Promise<StoredBlob>;
  get(key: string): Promise<Readable>;
  delete(key: string): Promise<void>;
}

export class BlobTooLargeError extends Error {
  override readonly name = 'BlobTooLargeError';
}

export class BlobNotFoundError extends Error {
  override readonly name = 'BlobNotFoundError';
}

class HashAndCount extends Transform {
  readonly hash = createHash('sha256');
  size = 0;

  constructor(private readonly maxBytes: number) {
    super();
  }

  override _transform(chunk: Buffer, _encoding: BufferEncoding, callback: TransformCallback): void {
    this.size += chunk.length;
    if (this.size > this.maxBytes) {
      callback(new BlobTooLargeError(`blob exceeds ${this.maxBytes} bytes`));
      return;
    }
    this.hash.update(chunk);
    callback(null, chunk);
  }
}

const KEY = /^[A-Za-z0-9._-]+(\/[A-Za-z0-9._-]+)*$/;

/** Stores blobs at `<root>/<key>`, for example `data/blobs/<fileId>/<versionId>`. */
export class LocalDiskBlobStore implements BlobStore {
  private readonly root: string;

  constructor(root: string) {
    this.root = resolve(root);
  }

  private pathFor(key: string): string {
    if (!KEY.test(key) || key.split('/').some((part) => part === '.' || part === '..')) {
      throw new Error(`invalid blob key: ${key}`);
    }
    const path = join(this.root, key);
    if (!path.startsWith(this.root + sep)) throw new Error(`invalid blob key: ${key}`);
    return path;
  }

  async put(key: string, body: Readable, opts: { maxBytes?: number } = {}): Promise<StoredBlob> {
    const path = this.pathFor(key);
    await mkdir(dirname(path), { recursive: true });
    if (await exists(path)) throw new Error(`blob ${key} already exists; blobs are immutable`);

    const tmp = `${path}.tmp-${randomUUID()}`;
    const meter = new HashAndCount(opts.maxBytes ?? Number.POSITIVE_INFINITY);
    try {
      await pipeline(body, meter, createWriteStream(tmp, { flags: 'wx' }));
      await rename(tmp, path);
    } catch (err) {
      await rm(tmp, { force: true });
      throw err;
    }
    return { key, size: meter.size, sha256: meter.hash.digest('hex') };
  }

  async get(key: string): Promise<Readable> {
    const path = this.pathFor(key);
    if (!(await exists(path))) throw new BlobNotFoundError(`blob ${key} not found`);
    return createReadStream(path);
  }

  async delete(key: string): Promise<void> {
    await rm(this.pathFor(key), { force: true });
  }
}

async function exists(path: string): Promise<boolean> {
  try {
    await access(path);
    return true;
  } catch {
    return false;
  }
}
