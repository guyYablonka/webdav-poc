import type { FastifyPluginAsync } from 'fastify';
import { loginPage } from '../ui/pages.js';
import type { MockOidcAuth } from './mock-oidc-auth.js';

/** Login and logout pages for the mock. An OIDC implementation would register its redirect and callback routes instead. */
export function authRoutes(auth: MockOidcAuth): FastifyPluginAsync {
  return async (app) => {
    app.get('/login', async (req, reply) => {
      if (auth.getCurrentUser(req)) return reply.redirect('/');
      return reply.type('text/html; charset=utf-8').send(loginPage(auth.users));
    });

    app.post<{ Body: { userId?: string } }>('/login', async (req, reply) => {
      const userId = req.body?.userId;
      if (!userId || !auth.users.some((u) => u.id === userId)) return reply.code(400).send({ error: 'unknown user' });
      auth.login(reply, userId);
      return reply.redirect('/', 303);
    });

    app.post('/logout', async (req, reply) => {
      const user = auth.logout(req, reply);
      if (user) req.log.info({ userId: user.id }, 'logged out; DAV sessions revoked');
      return reply.redirect('/login', 303);
    });
  };
}
