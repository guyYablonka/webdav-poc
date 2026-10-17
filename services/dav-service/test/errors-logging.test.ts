import { describe, expect, it } from 'vitest';
import { GATEWAY_UNAVAILABLE, GatewayError, davStatusForGatewayError } from '../src/gateway/errors.js';
import { redactHeaders, redactUrl } from '../src/logging.js';

describe('gateway error mapping', () => {
  const status = (s: number, body: GatewayError['body'] = null) => davStatusForGatewayError(new GatewayError(s, body, 'x'));

  it('follows the mapping table', () => {
    expect(status(401, { error: 'session_invalid', reason: 'unknown' })).toBe(401);
    expect(status(401, { error: 'session_invalid', reason: 'revoked' })).toBe(401);
    expect(status(401, { error: 'session_invalid', reason: 'beyond_cap' })).toBe(401);
    expect(status(403)).toBe(403);
    expect(status(412)).toBe(412);
    expect(status(423)).toBe(423);
    expect(status(413)).toBe(413);
    expect(status(GATEWAY_UNAVAILABLE)).toBe(503);
    expect(status(500)).toBe(503);
    expect(status(502)).toBe(503);
  });

  it('passes through the other statuses the gateway uses and flags the rest', () => {
    expect(status(404)).toBe(404);
    expect(status(409)).toBe(409);
    expect(status(400)).toBe(400);
    expect(status(418)).toBe(500);
  });
});

describe('log redaction', () => {
  const token = 'Zm9vYmFyYmF6cXV4LXNlY3JldC10b2tlbi12YWx1ZQ';

  it('removes the token from URLs', () => {
    expect(redactUrl(`/s/${token}/f1/a.docx`)).toBe('/s/[token]/f1/a.docx');
    expect(redactUrl(`/s/${token}/`)).toBe('/s/[token]/');
    expect(redactUrl('/healthz')).toBe('/healthz');
  });

  it('removes the token from header values and hides credentials', () => {
    const headers = redactHeaders({
      if: `<https://localhost:8443/s/${token}/f1/a.docx> (<opaquelocktoken:abc>)`,
      destination: [`https://localhost:8443/s/${token}/f1/b.docx`],
      authorization: 'Basic abc',
      cookie: 'a=b',
      depth: '0',
    });
    expect(JSON.stringify(headers)).not.toContain(token);
    expect(headers.if).toBe('<https://localhost:8443/s/[token]/f1/a.docx> (<opaquelocktoken:abc>)');
    expect(headers.authorization).toBe('[redacted]');
    expect(headers.cookie).toBe('[redacted]');
    expect(headers.depth).toBe('0');
  });
});
