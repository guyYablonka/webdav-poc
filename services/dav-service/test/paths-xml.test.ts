import { describe, expect, it } from 'vitest';
import { parseTarget, sameFileName } from '../src/dav/paths.js';
import { lockResponseBody, propfindBody, type DavResource } from '../src/dav/props.js';
import { XmlError, parseLockinfo, parsePropfind, parseProppatch } from '../src/dav/xml.js';

describe('URL space', () => {
  it('maps the parent collections Office probes', () => {
    expect(parseTarget('/')).toEqual({ kind: 'root', path: '/' });
    expect(parseTarget('/s/')).toEqual({ kind: 'root', path: '/s/' });
    expect(parseTarget('/s/tok/')).toEqual({ kind: 'session', token: 'tok' });
    expect(parseTarget('/s/tok/f1/')).toEqual({ kind: 'collection', token: 'tok', fileId: 'f1' });
    expect(parseTarget('/s/tok/f1')).toEqual({ kind: 'collection', token: 'tok', fileId: 'f1' });
  });

  it('decodes UTF-8 file names', () => {
    const name = 'דוח רבעוני (טיוטה).docx';
    const target = parseTarget(`/s/tok/f1/${encodeURIComponent(name)}?x=1`);
    expect(target).toEqual({ kind: 'file', token: 'tok', fileId: 'f1', fileName: name });
  });

  it('flags SharePoint and FrontPage probes', () => {
    expect(parseTarget('/_vti_inf.html').kind).toBe('office-probe');
    expect(parseTarget('/_vti_bin/shtml.dll/_vti_rpc').kind).toBe('office-probe');
    expect(parseTarget('/s/tok/f1/_vti_bin/cellstorage.svc').kind).toBe('office-probe');
  });

  it('rejects anything else', () => {
    expect(parseTarget('/other').kind).toBe('not-found');
    expect(parseTarget('/s/tok/f1/a.docx/extra').kind).toBe('not-found');
    expect(parseTarget('/s/tok/f1/a%2Fb.docx').kind).toBe('not-found');
    expect(parseTarget('/s/tok/f1/%E0%A4%A').kind).toBe('bad-request');
  });

  it('accepts absolute-form request targets', () => {
    expect(parseTarget('https://localhost:8443/s/tok/f1/a.docx')).toMatchObject({ kind: 'file', fileName: 'a.docx' });
  });

  it('compares names case- and normalisation-insensitively', () => {
    expect(sameFileName('Report.DOCX', 'report.docx')).toBe(true);
    expect(sameFileName('é.docx', 'é.docx')).toBe(true);
    expect(sameFileName('a.docx', 'b.docx')).toBe(false);
  });
});

describe('PROPFIND bodies', () => {
  it('treats an empty body as allprop', () => {
    expect(parsePropfind('')).toEqual({ kind: 'allprop' });
    expect(parsePropfind('  \n')).toEqual({ kind: 'allprop' });
  });

  it('reads requested properties with their namespaces', () => {
    const body = `<?xml version="1.0" encoding="utf-8"?>
      <D:propfind xmlns:D="DAV:" xmlns:Z="urn:schemas-microsoft-com:">
        <D:prop><D:getetag/><D:lockdiscovery/><Z:Win32FileAttributes/></D:prop>
      </D:propfind>`;
    expect(parsePropfind(body)).toEqual({
      kind: 'prop',
      props: [
        { ns: 'DAV:', local: 'getetag' },
        { ns: 'DAV:', local: 'lockdiscovery' },
        { ns: 'urn:schemas-microsoft-com:', local: 'Win32FileAttributes' },
      ],
    });
  });

  it('handles allprop, propname and a default namespace', () => {
    expect(parsePropfind('<propfind xmlns="DAV:"><allprop/></propfind>')).toEqual({ kind: 'allprop' });
    expect(parsePropfind('<a:propfind xmlns:a="DAV:"><a:propname/></a:propfind>')).toEqual({ kind: 'propname' });
  });

  it('rejects malformed or foreign XML', () => {
    expect(() => parsePropfind('<D:propfind xmlns:D="DAV:"><D:prop>')).toThrow(XmlError);
    expect(() => parsePropfind('<propfind/>')).toThrow(XmlError);
    expect(() =>
      parsePropfind('<!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/passwd">]><D:propfind xmlns:D="DAV:">&e;</D:propfind>'),
    ).toThrow(XmlError);
  });

  const file: DavResource = {
    href: 'https://localhost:8443/s/tok/f1/%D7%90%20(1).docx',
    isCollection: false,
    displayName: 'א (1).docx',
    createdAt: new Date('2026-01-01T10:00:00Z'),
    modifiedAt: new Date('2026-01-02T10:00:00Z'),
    file: { size: 42, mime: 'application/x-test', etag: '"v1"', lock: null },
  };

  it('renders found and missing properties in separate propstats', () => {
    const xml = propfindBody([file], {
      kind: 'prop',
      props: [
        { ns: 'DAV:', local: 'getcontentlength' },
        { ns: 'DAV:', local: 'nonsense' },
        { ns: 'urn:schemas-microsoft-com:', local: 'Win32FileAttributes' },
      ],
    });
    expect(xml).toContain('<D:getcontentlength>42</D:getcontentlength>');
    expect(xml).toMatch(/<D:nonsense\/>.*<x:Win32FileAttributes xmlns:x="urn:schemas-microsoft-com:"\/>.*404 Not Found/);
    expect(xml).toContain(`<D:href>${file.href}</D:href>`);
  });

  it('renders every live property for allprop with RFC 1123 dates', () => {
    const xml = propfindBody([file], { kind: 'allprop' });
    for (const p of ['displayname', 'getcontenttype', 'getetag', 'creationdate', 'resourcetype', 'supportedlock', 'lockdiscovery']) {
      expect(xml).toContain(`D:${p}`);
    }
    expect(xml).toContain('<D:getlastmodified>Fri, 02 Jan 2026 10:00:00 GMT</D:getlastmodified>');
    expect(xml).toContain('<D:creationdate>2026-01-01T10:00:00Z</D:creationdate>');
    expect(xml).toContain('<D:getetag>&quot;v1&quot;</D:getetag>');
    expect(xml).not.toContain('404');
  });
});

describe('LOCK bodies', () => {
  it('parses an exclusive write lock and keeps the owner element', () => {
    const info = parseLockinfo(`<?xml version="1.0" encoding="utf-8"?>
      <a:lockinfo xmlns:a="DAV:"><a:lockscope><a:exclusive/></a:lockscope><a:locktype><a:write/></a:locktype>
      <a:owner><a:href>CORP\\alice</a:href></a:owner></a:lockinfo>`);
    expect(info.scope).toBe('exclusive');
    expect(info.isWrite).toBe(true);
    expect(info.ownerXml).toBe('<a:owner xmlns:a="DAV:"><a:href>CORP\\alice</a:href></a:owner>');
  });

  it('reports shared locks', () => {
    const info = parseLockinfo(
      '<D:lockinfo xmlns:D="DAV:"><D:lockscope><D:shared/></D:lockscope><D:locktype><D:write/></D:locktype></D:lockinfo>',
    );
    expect(info.scope).toBe('shared');
    expect(info.ownerXml).toBeNull();
  });

  it('echoes the owner in the lock response', () => {
    const xml = lockResponseBody(
      { lockToken: 'opaquelocktoken:t', ownerXml: '<a:owner xmlns:a="DAV:">Alice</a:owner>', timeoutS: 3600 },
      'https://h/s/t/f/a.docx',
    );
    expect(xml).toContain('<a:owner xmlns:a="DAV:">Alice</a:owner>');
    expect(xml).toContain('<D:timeout>Second-3600</D:timeout>');
    expect(xml).toContain('<D:locktoken><D:href>opaquelocktoken:t</D:href></D:locktoken>');
  });

  it('drops an owner value that is not a well-formed DAV:owner element', () => {
    const xml = lockResponseBody({ lockToken: 'opaquelocktoken:t', ownerXml: '<oops>', timeoutS: 1 }, 'https://h/x');
    expect(xml).not.toContain('oops');
  });
});

describe('PROPPATCH bodies', () => {
  it('lists every property in set and remove', () => {
    const names = parseProppatch(`<D:propertyupdate xmlns:D="DAV:" xmlns:Z="urn:schemas-microsoft-com:">
      <D:set><D:prop><Z:Win32LastModifiedTime>x</Z:Win32LastModifiedTime></D:prop></D:set>
      <D:remove><D:prop><Z:Win32FileAttributes/></D:prop></D:remove></D:propertyupdate>`);
    expect(names.map((n) => n.local)).toEqual(['Win32LastModifiedTime', 'Win32FileAttributes']);
  });
});
