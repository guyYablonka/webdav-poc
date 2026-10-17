/**
 * Parsers for the WebDAV request headers dav-service relies on (RFC 4918).
 */

/**
 * Returns the state tokens of an `If` header in order, skipping negated ones
 * and ETag conditions. Handles both the untagged form `(<token>)` and the
 * tagged form `<url> (<token>)`; resource tags are ignored because every
 * session addresses exactly one file.
 */
export function parseIfHeader(value: string | undefined): string[] {
  if (!value) return [];
  const tokens: string[] = [];
  let i = 0;
  let inList = false;
  let negate = false;
  while (i < value.length) {
    const c = value[i]!;
    if (c === '(') {
      inList = true;
      negate = false;
      i++;
    } else if (c === ')') {
      inList = false;
      i++;
    } else if (c === '<') {
      const end = value.indexOf('>', i + 1);
      if (end === -1) break;
      const inner = value.slice(i + 1, end).trim();
      if (inList && !negate && inner) tokens.push(inner);
      negate = false;
      i = end + 1;
    } else if (c === '[') {
      // ETag condition; skip to the closing bracket.
      const end = value.indexOf(']', i + 1);
      if (end === -1) break;
      negate = false;
      i = end + 1;
    } else if (inList && value.slice(i, i + 3).toLowerCase() === 'not') {
      negate = true;
      i += 3;
    } else {
      i++;
    }
  }
  return tokens;
}

/** `Lock-Token: <opaquelocktoken:...>` -> `opaquelocktoken:...` */
export function parseLockTokenHeader(value: string | undefined): string | null {
  if (!value) return null;
  const trimmed = value.trim();
  const match = /^<([^>]+)>$/.exec(trimmed);
  const token = (match ? match[1]! : trimmed).trim();
  return token || null;
}

/**
 * `Timeout: Second-3600`, `Infinite, Second-4100000000`, ... The first value
 * we understand wins and is capped at `maxS`. No usable value yields `maxS`.
 */
export function parseTimeout(value: string | undefined, maxS: number): number {
  if (!value) return maxS;
  for (const part of value.split(',')) {
    const item = part.trim();
    if (/^infinite$/i.test(item)) return maxS;
    const match = /^second-(\d+)$/i.exec(item);
    if (match) {
      const seconds = Number(match[1]);
      if (Number.isFinite(seconds) && seconds > 0) return Math.min(seconds, maxS);
    }
  }
  return maxS;
}

export type Depth = '0' | '1' | 'infinity';

export function parseDepth(value: string | undefined, fallback: Depth): Depth | null {
  if (value === undefined) return fallback;
  const v = value.trim().toLowerCase();
  if (v === '0' || v === '1' || v === 'infinity') return v;
  return null;
}

/** True when an `If-None-Match` header matches the given strong ETag. */
export function etagMatches(header: string | undefined, etag: string): boolean {
  if (!header) return false;
  if (header.trim() === '*') return true;
  const wanted = etag.replace(/^W\//, '');
  return header.split(',').some((candidate) => candidate.trim().replace(/^W\//, '') === wanted);
}
