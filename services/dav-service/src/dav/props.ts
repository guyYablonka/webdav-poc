import { DOMParser, onErrorStopParsing } from '@xmldom/xmldom';
import type { FileLock } from '@office-dav/gateway-contract';
import {
  DAV_NS,
  XML_DECL,
  escapeXml,
  multistatus,
  propElement,
  type MultistatusResponse,
  type PropName,
  type PropfindRequest,
} from './xml.js';

/** What PROPFIND needs to know about one resource. */
export interface DavResource {
  href: string;
  isCollection: boolean;
  displayName: string;
  createdAt: Date;
  modifiedAt: Date;
  file?: {
    size: number;
    mime: string;
    /** Strong, quoted ETag. */
    etag: string;
    lock: FileLock | null;
  };
}

export interface ActiveLock {
  lockToken: string;
  ownerXml: string | null;
  timeoutS: number;
}

const SUPPORTED_LOCK =
  '<D:lockentry><D:lockscope><D:exclusive/></D:lockscope><D:locktype><D:write/></D:locktype></D:lockentry>';

/** RFC 3339 without milliseconds, which some clients parse more reliably. */
function rfc3339(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/**
 * The owner element is echoed back verbatim. It is checked to be one
 * well-formed DAV:owner element so a bad stored value cannot break PROPFIND.
 */
function safeOwnerXml(ownerXml: string | null): string {
  if (!ownerXml) return '';
  try {
    const root = new DOMParser({ onError: onErrorStopParsing }).parseFromString(ownerXml, 'application/xml')
      .documentElement;
    return root && root.namespaceURI === DAV_NS && root.localName === 'owner' ? ownerXml : '';
  } catch {
    return '';
  }
}

export function activeLockXml(lock: ActiveLock, lockRootHref: string): string {
  return (
    '<D:activelock>' +
    '<D:locktype><D:write/></D:locktype>' +
    '<D:lockscope><D:exclusive/></D:lockscope>' +
    '<D:depth>0</D:depth>' +
    safeOwnerXml(lock.ownerXml) +
    `<D:timeout>Second-${lock.timeoutS}</D:timeout>` +
    `<D:locktoken><D:href>${escapeXml(lock.lockToken)}</D:href></D:locktoken>` +
    `<D:lockroot><D:href>${escapeXml(lockRootHref)}</D:href></D:lockroot>` +
    '</D:activelock>'
  );
}

/** Body of a successful LOCK response. */
export function lockResponseBody(lock: ActiveLock, lockRootHref: string): string {
  return `${XML_DECL}<D:prop xmlns:D="DAV:"><D:lockdiscovery>${activeLockXml(lock, lockRootHref)}</D:lockdiscovery></D:prop>`;
}

function remainingSeconds(expiresAt: string, now: number): number {
  return Math.max(1, Math.ceil((Date.parse(expiresAt) - now) / 1000));
}

/** Inner XML of each live property, or null when it does not apply to the resource. */
type Renderer = (r: DavResource, now: number) => string | null;

const LIVE_PROPS: ReadonlyArray<readonly [string, Renderer]> = [
  ['displayname', (r) => escapeXml(r.displayName)],
  ['getcontentlength', (r) => (r.file ? String(r.file.size) : null)],
  ['getcontenttype', (r) => (r.file ? escapeXml(r.file.mime) : null)],
  ['getetag', (r) => (r.file ? escapeXml(r.file.etag) : null)],
  ['getlastmodified', (r) => r.modifiedAt.toUTCString()],
  ['creationdate', (r) => rfc3339(r.createdAt)],
  ['resourcetype', (r) => (r.isCollection ? '<D:collection/>' : '')],
  ['supportedlock', (r) => (r.isCollection ? '' : SUPPORTED_LOCK)],
  [
    'lockdiscovery',
    (r, now) => {
      const lock = r.file?.lock;
      if (!lock) return '';
      return activeLockXml({ ...lock, timeoutS: remainingSeconds(lock.expiresAt, now) }, r.href);
    },
  ],
];

const RENDERERS = new Map(LIVE_PROPS);

function renderResource(resource: DavResource, request: PropfindRequest, now: number): MultistatusResponse {
  const found: string[] = [];
  const missing: string[] = [];

  if (request.kind === 'prop') {
    for (const name of request.props) {
      const render = name.ns === DAV_NS ? RENDERERS.get(name.local) : undefined;
      const value = render ? render(resource, now) : null;
      if (value === null) missing.push(propElement(name));
      else found.push(propElement(name, value));
    }
  } else {
    for (const [local, render] of LIVE_PROPS) {
      const value = render(resource, now);
      if (value === null) continue;
      const name: PropName = { ns: DAV_NS, local };
      found.push(request.kind === 'propname' ? propElement(name) : propElement(name, value));
    }
  }

  return {
    href: resource.href,
    propstats: [
      { status: 200, props: found },
      { status: 404, props: missing },
    ],
  };
}

export function propfindBody(resources: DavResource[], request: PropfindRequest, now: number = Date.now()): string {
  return multistatus(resources.map((r) => renderResource(r, request, now)));
}

/** PROPPATCH is accepted and reported as successful, but nothing is stored. */
export function proppatchBody(href: string, names: PropName[]): string {
  return multistatus([{ href, propstats: [{ status: 200, props: names.map((n) => propElement(n)) }] }]);
}
