import { DOMParser, XMLSerializer, onErrorStopParsing, type Element, type Node } from '@xmldom/xmldom';

export const DAV_NS = 'DAV:';
export const XML_DECL = '<?xml version="1.0" encoding="utf-8"?>';

export class XmlError extends Error {
  override readonly name = 'XmlError';
}

/** A property name: namespace URI plus local name. */
export interface PropName {
  ns: string;
  local: string;
}

export function propKey(name: PropName): string {
  return `{${name.ns}}${name.local}`;
}

export function escapeXml(value: string): string {
  return value.replace(/[&<>"']/g, (c) =>
    c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : c === '"' ? '&quot;' : '&apos;',
  );
}

const ELEMENT_NODE = 1;

function parseRoot(body: string, expectedLocal: string): Element {
  let doc;
  try {
    // xmldom never fetches external entities or DTDs; unknown entities are errors.
    doc = new DOMParser({ onError: onErrorStopParsing }).parseFromString(body, 'application/xml');
  } catch (err) {
    throw new XmlError(`invalid XML: ${(err as Error).message}`);
  }
  const root = doc.documentElement;
  if (!root || root.namespaceURI !== DAV_NS || root.localName !== expectedLocal) {
    throw new XmlError(`expected a DAV:${expectedLocal} document`);
  }
  return root;
}

function childElements(node: Node): Element[] {
  const out: Element[] = [];
  for (let child = node.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === ELEMENT_NODE) out.push(child as Element);
  }
  return out;
}

function davChild(node: Node, local: string): Element | undefined {
  return childElements(node).find((el) => el.namespaceURI === DAV_NS && el.localName === local);
}

function nameOf(el: Element): PropName {
  return { ns: el.namespaceURI ?? '', local: el.localName ?? el.nodeName };
}

// ---------------------------------------------------------------------------
// PROPFIND
// ---------------------------------------------------------------------------

export type PropfindRequest = { kind: 'allprop' } | { kind: 'propname' } | { kind: 'prop'; props: PropName[] };

/** An empty body means allprop (RFC 4918, 9.1). */
export function parsePropfind(body: string): PropfindRequest {
  if (body.trim() === '') return { kind: 'allprop' };
  const root = parseRoot(body, 'propfind');
  if (davChild(root, 'propname')) return { kind: 'propname' };
  const prop = davChild(root, 'prop');
  if (prop) return { kind: 'prop', props: childElements(prop).map(nameOf) };
  if (davChild(root, 'allprop')) return { kind: 'allprop' };
  throw new XmlError('propfind needs allprop, propname or prop');
}

// ---------------------------------------------------------------------------
// PROPPATCH
// ---------------------------------------------------------------------------

/** Every property named in a propertyupdate's set and remove instructions. */
export function parseProppatch(body: string): PropName[] {
  const root = parseRoot(body, 'propertyupdate');
  const names: PropName[] = [];
  for (const instruction of childElements(root)) {
    if (instruction.namespaceURI !== DAV_NS) continue;
    if (instruction.localName !== 'set' && instruction.localName !== 'remove') continue;
    const prop = davChild(instruction, 'prop');
    if (prop) names.push(...childElements(prop).map(nameOf));
  }
  return names;
}

// ---------------------------------------------------------------------------
// LOCK
// ---------------------------------------------------------------------------

export interface LockInfo {
  scope: 'exclusive' | 'shared';
  isWrite: boolean;
  /** The serialised `<owner>` element, namespace declarations included, or null. */
  ownerXml: string | null;
}

export function parseLockinfo(body: string): LockInfo {
  const root = parseRoot(body, 'lockinfo');
  const scopeEl = davChild(root, 'lockscope');
  const scopeChild = scopeEl ? childElements(scopeEl)[0] : undefined;
  if (!scopeChild || scopeChild.namespaceURI !== DAV_NS) throw new XmlError('lockinfo needs a lockscope');
  if (scopeChild.localName !== 'exclusive' && scopeChild.localName !== 'shared') {
    throw new XmlError('unknown lockscope');
  }
  const typeEl = davChild(root, 'locktype');
  const typeChild = typeEl ? childElements(typeEl)[0] : undefined;
  const owner = davChild(root, 'owner');
  return {
    scope: scopeChild.localName,
    isWrite: typeChild?.namespaceURI === DAV_NS && typeChild.localName === 'write',
    ownerXml: owner ? new XMLSerializer().serializeToString(owner) : null,
  };
}

// ---------------------------------------------------------------------------
// Responses
// ---------------------------------------------------------------------------

const STATUS_TEXT: Record<number, string> = {
  200: 'OK',
  403: 'Forbidden',
  404: 'Not Found',
  424: 'Failed Dependency',
};

export function statusLine(status: number): string {
  return `HTTP/1.1 ${status} ${STATUS_TEXT[status] ?? ''}`.trimEnd();
}

/**
 * Renders one property element. DAV: properties use the `D:` prefix declared
 * on the multistatus root; others declare their namespace inline.
 */
export function propElement(name: PropName, innerXml = ''): string {
  const local = escapeXml(name.local);
  if (name.ns === DAV_NS) return innerXml ? `<D:${local}>${innerXml}</D:${local}>` : `<D:${local}/>`;
  const decl = `xmlns:x="${escapeXml(name.ns)}"`;
  if (name.ns === '') return innerXml ? `<${local} xmlns="">${innerXml}</${local}>` : `<${local} xmlns=""/>`;
  return innerXml ? `<x:${local} ${decl}>${innerXml}</x:${local}>` : `<x:${local} ${decl}/>`;
}

export interface Propstat {
  status: number;
  /** Rendered property elements. */
  props: string[];
}

export interface MultistatusResponse {
  href: string;
  propstats: Propstat[];
}

export function multistatus(responses: MultistatusResponse[]): string {
  const body = responses
    .map((r) => {
      const propstats = r.propstats
        .filter((ps) => ps.props.length > 0)
        .map(
          (ps) =>
            `<D:propstat><D:prop>${ps.props.join('')}</D:prop><D:status>${statusLine(ps.status)}</D:status></D:propstat>`,
        )
        .join('');
      return `<D:response><D:href>${escapeXml(r.href)}</D:href>${propstats}</D:response>`;
    })
    .join('');
  return `${XML_DECL}<D:multistatus xmlns:D="DAV:">${body}</D:multistatus>`;
}

export function davError(condition: string): string {
  return `${XML_DECL}<D:error xmlns:D="DAV:"><D:${condition}/></D:error>`;
}
