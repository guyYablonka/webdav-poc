/**
 * Maps a request URL onto the dav-service URL space:
 *
 *   /                              minimal collection (Office probes parents)
 *   /s/                            minimal collection
 *   /s/:token/                     minimal collection for the session
 *   /s/:token/:fileId/             collection holding only the session's file
 *   /s/:token/:fileId/:fileName    the file
 *
 * Office also probes for SharePoint and FrontPage (`_vti_bin/*`,
 * `_vti_inf.html`). Those get a fast 404 without touching the gateway.
 */
export type DavTarget =
  | { kind: 'root'; path: '/' | '/s/' }
  | { kind: 'session'; token: string }
  | { kind: 'collection'; token: string; fileId: string }
  | { kind: 'file'; token: string; fileId: string; fileName: string }
  | { kind: 'office-probe' }
  | { kind: 'not-found' }
  | { kind: 'bad-request' };

export function parseTarget(rawUrl: string): DavTarget {
  let path = rawUrl;
  // Absolute-form request targets (`https://host/path`) can arrive through proxies.
  if (!path.startsWith('/')) {
    try {
      path = new URL(path).pathname;
    } catch {
      return { kind: 'bad-request' };
    }
  }
  const cut = path.search(/[?#]/);
  if (cut !== -1) path = path.slice(0, cut);

  if (/\/_vti_/i.test(path)) return { kind: 'office-probe' };

  const segments: string[] = [];
  for (const raw of path.split('/')) {
    if (raw === '') continue;
    let decoded: string;
    try {
      decoded = decodeURIComponent(raw);
    } catch {
      return { kind: 'bad-request' };
    }
    // An encoded slash or NUL inside a segment never names one of our resources.
    if (decoded.includes('/') || decoded.includes('\0')) return { kind: 'not-found' };
    segments.push(decoded);
  }

  if (segments.length === 0) return { kind: 'root', path: '/' };
  if (segments[0] !== 's') return { kind: 'not-found' };
  const [, token, fileId, fileName, ...rest] = segments;
  if (rest.length > 0) return { kind: 'not-found' };
  if (token === undefined) return { kind: 'root', path: '/s/' };
  if (fileId === undefined) return { kind: 'session', token };
  if (fileName === undefined) return { kind: 'collection', token, fileId };
  return { kind: 'file', token, fileId, fileName };
}

/** Office may change case or Unicode normalisation of the name it was given. */
export function sameFileName(a: string, b: string): boolean {
  return a.normalize('NFC').toLowerCase() === b.normalize('NFC').toLowerCase();
}
