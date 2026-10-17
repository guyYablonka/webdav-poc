import { pino, stdSerializers, type DestinationStream, type Logger } from 'pino';

/**
 * Session tokens travel in the URL path (`/s/<token>/...`) and in `If` headers
 * that carry the full URL. Every log line goes through these helpers so the
 * token never reaches the logs; the session id is logged instead.
 */
const TOKEN_IN_PATH = /\/s\/[^/?#\s<>"]+/g;
const SECRET_HEADERS = new Set(['authorization', 'proxy-authorization', 'cookie', 'set-cookie']);

export function redactUrl(value: string): string {
  return value.replace(TOKEN_IN_PATH, '/s/[token]');
}

export function redactHeaders(headers: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [name, value] of Object.entries(headers)) {
    if (value === undefined) continue;
    if (SECRET_HEADERS.has(name.toLowerCase())) {
      out[name] = '[redacted]';
    } else if (Array.isArray(value)) {
      out[name] = value.map((v) => redactUrl(String(v)));
    } else {
      out[name] = redactUrl(String(value));
    }
  }
  return out;
}

interface LoggableRequest {
  method?: string;
  url?: string;
  headers?: Record<string, unknown>;
}

export function createLogger(level: string, destination?: DestinationStream): Logger {
  const options = {
    level,
    base: { service: 'dav-service' },
    serializers: {
      // Fastify passes its request and reply objects under these keys when it logs errors.
      req: (req: LoggableRequest) => ({ method: req.method, url: req.url ? redactUrl(req.url) : undefined }),
      res: (res: { statusCode?: number }) => ({ statusCode: res.statusCode }),
      err: (err: Error) => {
        const serialized = stdSerializers.err(err);
        serialized.message = redactUrl(serialized.message ?? '');
        if (serialized.stack) serialized.stack = redactUrl(serialized.stack);
        return serialized;
      },
    },
  };
  return destination ? pino(options, destination) : pino(options);
}
