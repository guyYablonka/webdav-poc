import { z } from 'zod';

/**
 * Gateway API v1.
 *
 * dav-service talks to exactly one thing: a service that implements these
 * endpoints under GATEWAY_BASE_PATH. The demo site implements them for the POC;
 * the production web app will implement the same contract.
 *
 * All timestamps are ISO 8601 strings in UTC.
 */
export const GATEWAY_BASE_PATH = '/gw/v1';

/** Header names, lower-cased as Node and undici expose them. */
export const GatewayHeaders = {
  /** Correlates a dav-service request with the gateway calls it makes. */
  requestId: 'x-request-id',
  /** The DAV session id returned by resolve. Sent on every call once known. */
  davSession: 'x-dav-session',
  /** The lock token a PUT is made under. */
  lockToken: 'x-lock-token',
  /** The version id of a content stream (same value as the unquoted ETag). */
  versionId: 'x-version-id',
} as const;

/** Endpoint paths, relative to GATEWAY_BASE_PATH. */
export const GatewayPaths = {
  resolve: '/sessions/resolve',
  file: (fileId: string) => `/files/${encodeURIComponent(fileId)}`,
  content: (fileId: string) => `/files/${encodeURIComponent(fileId)}/content`,
  locks: '/locks',
  lock: (lockToken: string) => `/locks/${encodeURIComponent(lockToken)}`,
} as const;

const Id = z.string().min(1).max(200);
const IsoDateTime = z.iso.datetime({ offset: true });

// ---------------------------------------------------------------------------
// POST /sessions/resolve
// ---------------------------------------------------------------------------

/** What the caller is about to do with the session. Lets the gateway audit or narrow access. */
export const Intent = z.enum(['read', 'write', 'lock']);
export type Intent = z.infer<typeof Intent>;

export const ResolveRequest = z.object({
  token: z.string().min(1).max(512),
  intent: Intent,
});
export type ResolveRequest = z.infer<typeof ResolveRequest>;

/**
 * `active`: the session is usable and any lock it holds is current.
 * `expired`: the session is usable but the lock it holds has lapsed. The next
 * LOCK refresh or PUT with that lock token may revive it. dav-service treats
 * both states as authenticated; the gateway decides on revival.
 */
export const SessionState = z.enum(['active', 'expired']);
export type SessionState = z.infer<typeof SessionState>;

export const ResolveResponse = z.object({
  sessionId: Id,
  userId: Id,
  displayName: z.string(),
  fileId: Id,
  canWrite: z.boolean(),
  state: SessionState,
});
export type ResolveResponse = z.infer<typeof ResolveResponse>;

/** Why a token does not authenticate. Every reason maps to WebDAV 401. */
export const SessionFailureReason = z.enum(['unknown', 'revoked', 'beyond_cap']);
export type SessionFailureReason = z.infer<typeof SessionFailureReason>;

// ---------------------------------------------------------------------------
// GET /files/:fileId
// ---------------------------------------------------------------------------

export const FileLock = z.object({
  /** Full lock token URI, for example `opaquelocktoken:6f1c...`. */
  lockToken: z.string().min(1).max(200),
  /** The `<D:owner>` element the client sent, echoed back verbatim. */
  ownerXml: z.string().nullable(),
  expiresAt: IsoDateTime,
  timeoutS: z.number().int().positive(),
  sessionId: Id,
});
export type FileLock = z.infer<typeof FileLock>;

export const FileInfo = z.object({
  fileId: Id,
  name: z.string().min(1),
  mime: z.string().min(1),
  size: z.number().int().nonnegative(),
  versionId: Id,
  updatedAt: IsoDateTime,
  createdAt: IsoDateTime,
  /** The active lock, or null when the file is unlocked or the lock has lapsed. */
  lock: FileLock.nullable(),
});
export type FileInfo = z.infer<typeof FileInfo>;

// ---------------------------------------------------------------------------
// PUT /files/:fileId/content
// ---------------------------------------------------------------------------

export const PutContentResponse = z.object({
  versionId: Id,
  size: z.number().int().nonnegative(),
  updatedAt: IsoDateTime,
  /** True when the PUT revived a lapsed lock (see the lock policy). */
  revived: z.boolean(),
});
export type PutContentResponse = z.infer<typeof PutContentResponse>;

// ---------------------------------------------------------------------------
// POST /locks, DELETE /locks/:lockToken
// ---------------------------------------------------------------------------

/** Without `lockToken`: acquire. With `lockToken`: refresh, or revive a lapsed lock. */
export const LockRequest = z.object({
  fileId: Id,
  sessionId: Id,
  timeoutS: z.number().int().positive(),
  ownerXml: z.string().max(8192).optional(),
  lockToken: z.string().min(1).max(200).optional(),
});
export type LockRequest = z.infer<typeof LockRequest>;

export const LockResponse = z.object({
  lockToken: z.string().min(1),
  expiresAt: IsoDateTime,
  /** The granted timeout, which may be shorter than requested. */
  timeoutS: z.number().int().positive(),
  revived: z.boolean(),
});
export type LockResponse = z.infer<typeof LockResponse>;

export const LockHolder = z.object({
  userId: Id,
  displayName: z.string(),
  since: IsoDateTime,
});
export type LockHolder = z.infer<typeof LockHolder>;

// ---------------------------------------------------------------------------
// Errors
// ---------------------------------------------------------------------------

export const GatewayErrorCode = z.enum([
  /** The caller's service credentials were rejected. */
  'gateway_auth',
  /** The DAV session does not authenticate; see `reason`. */
  'session_invalid',
  'forbidden',
  'not_found',
  'conflict',
  'precondition_failed',
  'payload_too_large',
  'locked',
  'bad_request',
  'internal',
]);
export type GatewayErrorCode = z.infer<typeof GatewayErrorCode>;

/**
 * Body of every non-2xx gateway response. `reason` is set on session 401s,
 * `heldBy` on 423s when another session holds the lock.
 */
export const GatewayErrorBody = z.object({
  error: GatewayErrorCode,
  message: z.string().optional(),
  reason: SessionFailureReason.optional(),
  heldBy: LockHolder.optional(),
});
export type GatewayErrorBody = z.infer<typeof GatewayErrorBody>;
