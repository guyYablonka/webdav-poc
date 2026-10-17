import type { GatewayErrorBody } from '@office-dav/gateway-contract';

/** Status used when the gateway could not be reached or returned something unparseable. */
export const GATEWAY_UNAVAILABLE = 0;

export class GatewayError extends Error {
  override readonly name = 'GatewayError';

  constructor(
    /** HTTP status from the gateway, or GATEWAY_UNAVAILABLE. */
    readonly status: number,
    readonly body: GatewayErrorBody | null,
    message: string,
    options?: { cause?: unknown },
  ) {
    super(message, options);
  }
}

/**
 * The one place that turns gateway outcomes into WebDAV statuses.
 *
 *   401 (any reason: unknown, revoked, beyond_cap) -> 401
 *   403 -> 403, 404 -> 404, 409 -> 409, 412 -> 412, 413 -> 413, 423 -> 423
 *   400 -> 400 (the client sent something the gateway rejected)
 *   unreachable, unparseable, or 5xx -> 503
 *   any other status -> 500 (a contract mismatch, so it is our bug)
 */
export function davStatusForGatewayError(err: GatewayError): number {
  if (err.status === GATEWAY_UNAVAILABLE || err.status >= 500) return 503;
  switch (err.status) {
    case 400:
    case 401:
    case 403:
    case 404:
    case 409:
    case 412:
    case 413:
    case 423:
      return err.status;
    default:
      return 500;
  }
}
