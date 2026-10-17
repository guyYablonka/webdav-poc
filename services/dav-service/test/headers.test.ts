import { describe, expect, it } from 'vitest';
import { etagMatches, parseDepth, parseIfHeader, parseLockTokenHeader, parseTimeout } from '../src/dav/headers.js';

describe('If header', () => {
  it('reads the untagged form', () => {
    expect(parseIfHeader('(<opaquelocktoken:abc-123>)')).toEqual(['opaquelocktoken:abc-123']);
  });

  it('reads the tagged form and ignores the resource tag', () => {
    const header = '<https://localhost:8443/s/tok/f1/A%20B.docx> (<opaquelocktoken:abc>)';
    expect(parseIfHeader(header)).toEqual(['opaquelocktoken:abc']);
  });

  it('skips ETags and negated tokens', () => {
    expect(parseIfHeader('(<opaquelocktoken:a> ["etag-1"]) (Not <DAV:no-lock>)')).toEqual(['opaquelocktoken:a']);
    expect(parseIfHeader('(["etag"] <opaquelocktoken:b>)')).toEqual(['opaquelocktoken:b']);
  });

  it('handles several lists and missing input', () => {
    expect(parseIfHeader('(<opaquelocktoken:a>) (<opaquelocktoken:b>)')).toEqual(['opaquelocktoken:a', 'opaquelocktoken:b']);
    expect(parseIfHeader(undefined)).toEqual([]);
    expect(parseIfHeader('garbage')).toEqual([]);
    expect(parseIfHeader('(<unterminated')).toEqual([]);
  });
});

describe('Lock-Token header', () => {
  it('strips angle brackets', () => {
    expect(parseLockTokenHeader('<opaquelocktoken:abc>')).toBe('opaquelocktoken:abc');
    expect(parseLockTokenHeader(' opaquelocktoken:abc ')).toBe('opaquelocktoken:abc');
    expect(parseLockTokenHeader('')).toBeNull();
    expect(parseLockTokenHeader(undefined)).toBeNull();
  });
});

describe('Timeout header', () => {
  it('honours Second-N up to the cap', () => {
    expect(parseTimeout('Second-600', 3600)).toBe(600);
    expect(parseTimeout('Second-7200', 3600)).toBe(3600);
  });

  it('maps Infinite and missing values to the cap', () => {
    expect(parseTimeout('Infinite, Second-4100000000', 3600)).toBe(3600);
    expect(parseTimeout(undefined, 120)).toBe(120);
    expect(parseTimeout('bogus', 120)).toBe(120);
  });

  it('takes the first value it understands', () => {
    expect(parseTimeout('Fortnight, Second-30', 3600)).toBe(30);
  });
});

describe('Depth and ETag helpers', () => {
  it('parses Depth', () => {
    expect(parseDepth('0', 'infinity')).toBe('0');
    expect(parseDepth('Infinity', '0')).toBe('infinity');
    expect(parseDepth(undefined, '1')).toBe('1');
    expect(parseDepth('2', '0')).toBeNull();
  });

  it('matches ETags', () => {
    expect(etagMatches('"v1"', '"v1"')).toBe(true);
    expect(etagMatches('"v0", W/"v1"', '"v1"')).toBe(true);
    expect(etagMatches('*', '"v1"')).toBe(true);
    expect(etagMatches('"v2"', '"v1"')).toBe(false);
    expect(etagMatches(undefined, '"v1"')).toBe(false);
  });
});
