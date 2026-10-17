import { createHash, timingSafeEqual } from 'node:crypto';
import type { FastifyRequest } from 'fastify';

/**
 * Decides whether a gateway call comes from a trusted service (dav-service).
 * The POC compares a shared bearer secret; production swaps in JWT
 * verification against dav-service's public key, or mTLS.
 */
export type ServiceAuthenticator = (req: FastifyRequest) => boolean;

export function sharedSecretAuth(secret: string): ServiceAuthenticator {
  const expected = createHash('sha256').update(secret).digest();
  return (req) => {
    const header = req.headers.authorization;
    if (!header?.startsWith('Bearer ')) return false;
    const presented = createHash('sha256').update(header.slice('Bearer '.length)).digest();
    return timingSafeEqual(presented, expected);
  };
}
