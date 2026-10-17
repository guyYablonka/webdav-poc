import type { FastifyReply, FastifyRequest } from 'fastify';
import type { Auth, AuthHooks, AuthUser } from './auth.js';

export const MOCK_USERS: readonly AuthUser[] = [
  { id: 'alice', displayName: 'Alice' },
  { id: 'bob', displayName: 'Bob' },
];

const COOKIE = 'demo_session';

/**
 * Stand-in for OIDC SSO: the user picks a name and gets a signed cookie
 * (`<userId>.<issuedAtMs>`, HMAC-signed by @fastify/cookie). Replace this
 * class with an OIDC implementation of `Auth`.
 */
export class MockOidcAuth implements Auth {
  constructor(
    private readonly hooks: AuthHooks,
    private readonly opts: { secureCookie: boolean; maxAgeS: number } = { secureCookie: true, maxAgeS: 12 * 3600 },
  ) {}

  get users(): readonly AuthUser[] {
    return MOCK_USERS;
  }

  getCurrentUser(req: FastifyRequest): AuthUser | null {
    const raw = req.cookies[COOKIE];
    if (!raw) return null;
    const unsigned = req.unsignCookie(raw);
    if (!unsigned.valid || !unsigned.value) return null;
    const [userId, issuedAt] = unsigned.value.split('.');
    if (!userId || !issuedAt || Date.now() - Number(issuedAt) > this.opts.maxAgeS * 1000) return null;
    return MOCK_USERS.find((u) => u.id === userId) ?? null;
  }

  login(reply: FastifyReply, userId: string): AuthUser {
    const user = MOCK_USERS.find((u) => u.id === userId);
    if (!user) throw new Error(`unknown user ${userId}`);
    reply.setCookie(COOKIE, `${user.id}.${Date.now()}`, {
      signed: true,
      httpOnly: true,
      secure: this.opts.secureCookie,
      sameSite: 'lax',
      path: '/',
      maxAge: this.opts.maxAgeS,
    });
    this.hooks.onLogin(user);
    return user;
  }

  logout(req: FastifyRequest, reply: FastifyReply): AuthUser | null {
    const user = this.getCurrentUser(req);
    reply.clearCookie(COOKIE, { path: '/' });
    if (user) this.hooks.onLogout(user);
    return user;
  }
}
