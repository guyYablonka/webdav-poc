import type { FastifyReply, FastifyRequest } from 'fastify';

export interface AuthUser {
  id: string;
  displayName: string;
}

/**
 * Browser authentication for the demo site. The POC uses MockOidcAuth; the
 * real app plugs in its OIDC SSO here. Nothing outside `auth/` knows how a
 * user signed in.
 */
export interface Auth {
  /** The signed-in user, or null. Synchronous and cheap: called on every page and API request. */
  getCurrentUser(req: FastifyRequest): AuthUser | null;
  /** Establishes the browser session once the user's identity is known. */
  login(reply: FastifyReply, userId: string): AuthUser;
  /** Ends the browser session. Implementations call `onLogout`, which revokes the user's DAV sessions. */
  logout(req: FastifyRequest, reply: FastifyReply): AuthUser | null;
}

export interface AuthHooks {
  /** JIT provisioning: keep the local users table in step with the identity provider. */
  onLogin(user: AuthUser): void;
  onLogout(user: AuthUser): void;
}
