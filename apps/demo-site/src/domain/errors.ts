import type { GatewayErrorCode, LockHolder, SessionFailureReason } from '@office-dav/gateway-contract';

/**
 * A rule refused the operation. `status` and the body fields map one-to-one
 * onto the Gateway API error responses.
 */
export class PolicyError extends Error {
  override readonly name = 'PolicyError';

  constructor(
    readonly status: number,
    readonly code: GatewayErrorCode,
    message: string,
    readonly details: { reason?: SessionFailureReason; heldBy?: LockHolder } = {},
  ) {
    super(message);
  }
}

export function sessionInvalid(reason: SessionFailureReason): PolicyError {
  return new PolicyError(401, 'session_invalid', `DAV session is ${reason}`, { reason });
}

export function notFound(what: string): PolicyError {
  return new PolicyError(404, 'not_found', `${what} not found`);
}
