import { describe, expect, it } from 'vitest';
import {
  FileInfo,
  GatewayPaths,
  LockRequest,
  ResolveResponse,
  davFileUrl,
  encodePathSegment,
  officeAppForFileName,
  officeEditUri,
} from '../src/index.js';

describe('dav URLs', () => {
  it('encodes Hebrew, spaces and parentheses in the file name', () => {
    const name = 'תוכנית עבודה (טיוטה).docx';
    const url = davFileUrl('https://localhost:8443/', 'tok', 'f1', name);
    expect(url).toMatch(/^https:\/\/localhost:8443\/s\/tok\/f1\/[A-Za-z0-9%._~-]+$/);
    expect(url).not.toContain(' ');
    expect(url).not.toContain('(');
    expect(decodeURIComponent(url.split('/').pop()!)).toBe(name);
  });

  it('leaves unreserved characters alone', () => {
    expect(encodePathSegment('a-b_c.d~e')).toBe('a-b_c.d~e');
    expect(encodePathSegment("x!y'z*")).toBe('x%21y%27z%2A');
  });

  it('builds Office URIs per application', () => {
    expect(officeAppForFileName('A.DOCX')).toBe('word');
    expect(officeAppForFileName('b.xlsx')).toBe('excel');
    expect(officeAppForFileName('c.pptx')).toBe('powerpoint');
    expect(officeAppForFileName('d.pdf')).toBeNull();
    expect(officeEditUri('excel', 'https://h/s/t/f/b.xlsx')).toBe('ms-excel:ofe|u|https://h/s/t/f/b.xlsx');
  });

  it('encodes lock tokens in gateway paths', () => {
    expect(GatewayPaths.lock('opaquelocktoken:abc')).toBe('/locks/opaquelocktoken%3Aabc');
  });
});

describe('schemas', () => {
  it('accepts a well-formed file info', () => {
    const info = FileInfo.parse({
      fileId: 'f1',
      name: 'a.docx',
      mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      size: 10,
      versionId: 'v1',
      updatedAt: '2026-01-01T00:00:00.000Z',
      createdAt: '2026-01-01T00:00:00.000Z',
      lock: null,
    });
    expect(info.lock).toBeNull();
  });

  it('rejects an unknown session state', () => {
    expect(() =>
      ResolveResponse.parse({
        sessionId: 's',
        userId: 'u',
        displayName: 'U',
        fileId: 'f',
        canWrite: true,
        state: 'idle',
      }),
    ).toThrow();
  });

  it('distinguishes acquire from refresh by lockToken', () => {
    expect(LockRequest.parse({ fileId: 'f', sessionId: 's', timeoutS: 60 }).lockToken).toBeUndefined();
    expect(LockRequest.parse({ fileId: 'f', sessionId: 's', timeoutS: 60, lockToken: 'opaquelocktoken:x' }).lockToken).toBe(
      'opaquelocktoken:x',
    );
  });
});
