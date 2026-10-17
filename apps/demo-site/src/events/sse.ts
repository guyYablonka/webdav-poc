import type { ServerResponse } from 'node:http';
import type { FastifyPluginAsync } from 'fastify';
import type { Auth } from '../auth/auth.js';
import type { EventBus } from './bus.js';

/** `GET /events`: Server-Sent Events for signed-in browsers. */
export function sseRoutes(deps: { events: EventBus; auth: Auth }): FastifyPluginAsync {
  return async (app) => {
    const open = new Set<ServerResponse>();

    app.addHook('onClose', async () => {
      for (const res of open) res.end();
    });

    app.get('/events', (req, reply) => {
      if (!deps.auth.getCurrentUser(req)) return reply.code(401).send({ error: 'unauthorized' });

      reply.hijack();
      const res = reply.raw;
      res.writeHead(200, {
        'content-type': 'text/event-stream; charset=utf-8',
        'cache-control': 'no-cache, no-transform',
        connection: 'keep-alive',
        'x-accel-buffering': 'no',
      });
      res.write('retry: 3000\n\n');

      const unsubscribe = deps.events.subscribe((event) => {
        res.write(`event: ${event.type}\ndata: ${JSON.stringify(event)}\n\n`);
      });
      const ping = setInterval(() => res.write(': ping\n\n'), 15_000);
      open.add(res);

      req.raw.on('close', () => {
        clearInterval(ping);
        unsubscribe();
        open.delete(res);
      });
    });
  };
}
