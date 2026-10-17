import { EventEmitter } from 'node:events';

/** Events pushed to the browser over SSE. `at` is an ISO timestamp. */
export type DemoEvent =
  | {
      type: 'version.created';
      fileId: string;
      versionId: string;
      versionNo: number;
      userId: string;
      source: 'seed' | 'web' | 'dav';
      at: string;
    }
  | { type: 'lock.acquired'; fileId: string; userId: string; sessionId: string; expiresAt: string; at: string }
  | { type: 'lock.refreshed'; fileId: string; userId: string; sessionId: string; expiresAt: string; at: string }
  | { type: 'lock.revived'; fileId: string; userId: string; sessionId: string; expiresAt: string; at: string }
  | {
      type: 'lock.released';
      fileId: string;
      userId: string;
      reason: 'unlock' | 'force_unlock' | 'session_revoked';
      at: string;
    }
  | { type: 'file.created'; fileId: string; name: string; userId: string; at: string }
  | { type: 'session.revoked'; fileId: string; userId: string; sessionId: string; at: string };

export type DemoEventType = DemoEvent['type'];

export const EVENT_TYPES: readonly DemoEventType[] = [
  'version.created',
  'lock.acquired',
  'lock.refreshed',
  'lock.revived',
  'lock.released',
  'file.created',
  'session.revoked',
];

/** In-process pub/sub. A multi-instance deployment would back this with Redis or Postgres LISTEN/NOTIFY. */
export class EventBus {
  private readonly emitter = new EventEmitter();

  constructor() {
    this.emitter.setMaxListeners(0);
  }

  publish(event: DemoEvent): void {
    this.emitter.emit('event', event);
  }

  subscribe(listener: (event: DemoEvent) => void): () => void {
    this.emitter.on('event', listener);
    return () => this.emitter.off('event', listener);
  }
}
