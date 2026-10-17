import { createHash } from 'node:crypto';
import type { Intent, ResolveResponse } from '@office-dav/gateway-contract';
import type { CallContext, GatewayClient } from './client.js';

interface Entry {
  value: ResolveResponse;
  expiresAt: number;
}

/**
 * Resolves session tokens through the gateway, caching successful results per
 * token for `ttlMs` (RESOLVE_CACHE_MS). Concurrent lookups of the same token
 * share one gateway call. Failures are never cached.
 *
 * The cache only saves round trips: the gateway re-checks the session on every
 * call that carries X-Dav-Session, so a revocation takes effect immediately.
 */
export class SessionResolver {
  private readonly cache = new Map<string, Entry>();
  private readonly inflight = new Map<string, Promise<ResolveResponse>>();

  constructor(
    private readonly gateway: GatewayClient,
    private readonly ttlMs: number,
    private readonly maxEntries = 10_000,
    private readonly now: () => number = Date.now,
  ) {}

  async resolve(token: string, intent: Intent, ctx: CallContext): Promise<ResolveResponse> {
    if (this.ttlMs <= 0) return this.gateway.resolve(token, intent, ctx);

    // Key by hash so raw tokens are not kept in memory longer than a request.
    const key = createHash('sha256').update(token).digest('base64url');
    const hit = this.cache.get(key);
    if (hit && hit.expiresAt > this.now()) return hit.value;

    let pending = this.inflight.get(key);
    if (!pending) {
      pending = this.gateway
        .resolve(token, intent, ctx)
        .then((value) => {
          this.store(key, value);
          return value;
        })
        .finally(() => this.inflight.delete(key));
      this.inflight.set(key, pending);
    }
    return pending;
  }

  invalidate(token: string): void {
    this.cache.delete(createHash('sha256').update(token).digest('base64url'));
  }

  private store(key: string, value: ResolveResponse): void {
    this.cache.delete(key);
    this.cache.set(key, { value, expiresAt: this.now() + this.ttlMs });
    while (this.cache.size > this.maxEntries) {
      const oldest = this.cache.keys().next().value;
      if (oldest === undefined) break;
      this.cache.delete(oldest);
    }
  }
}
