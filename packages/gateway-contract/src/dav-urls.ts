/**
 * The dav-service URL space, shared because the web app builds the URLs that
 * Office opens and dav-service has to parse them back.
 *
 *   /s/<token>/<fileId>/<fileName>   the file
 *   /s/<token>/<fileId>/             a collection holding only that file
 */

/**
 * Percent-encodes one path segment as UTF-8. Only RFC 3986 unreserved
 * characters stay literal, so spaces, parentheses and Hebrew are all encoded.
 */
export function encodePathSegment(segment: string): string {
  return encodeURIComponent(segment).replace(
    /[!'()*]/g,
    (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`,
  );
}

function trimSlash(baseUrl: string): string {
  return baseUrl.replace(/\/+$/, '');
}

export function davCollectionUrl(baseUrl: string, token: string, fileId: string): string {
  return `${trimSlash(baseUrl)}/s/${encodePathSegment(token)}/${encodePathSegment(fileId)}/`;
}

export function davFileUrl(baseUrl: string, token: string, fileId: string, fileName: string): string {
  return `${davCollectionUrl(baseUrl, token, fileId)}${encodePathSegment(fileName)}`;
}

export type OfficeApp = 'word' | 'excel' | 'powerpoint';

const OFFICE_APP_BY_EXTENSION: Record<string, OfficeApp> = {
  docx: 'word',
  xlsx: 'excel',
  pptx: 'powerpoint',
};

export function officeAppForFileName(fileName: string): OfficeApp | null {
  const ext = fileName.toLowerCase().split('.').pop() ?? '';
  return OFFICE_APP_BY_EXTENSION[ext] ?? null;
}

/**
 * Office URI scheme for "open for edit", for example
 * `ms-word:ofe|u|https://localhost:8443/s/<token>/<fileId>/Report.docx`.
 */
export function officeEditUri(app: OfficeApp, davUrl: string): string {
  return `ms-${app}:ofe|u|${davUrl}`;
}
