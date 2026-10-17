import type { AuthUser } from '../auth/auth.js';
import type { FileDetail, FileSummary } from '../domain/files.js';
import { formatBytes, html, raw, type SafeHtml } from './html.js';

/** Times are rendered as <time> and formatted in the browser's locale by app.js. */
function time(iso: string, fmt: 'time' | 'datetime' = 'datetime'): SafeHtml {
  return html`<time datetime="${iso}" data-fmt="${fmt}">${iso}</time>`;
}

function fileIcon(f: FileSummary): SafeHtml {
  const ext = f.type?.extension ?? 'file';
  const letter = ext === 'docx' ? 'W' : ext === 'xlsx' ? 'X' : ext === 'pptx' ? 'P' : '?';
  return html`<span class="icon ${ext}" aria-hidden="true">${letter}</span>`;
}

function lockBadge(f: FileSummary): SafeHtml {
  if (!f.lock) return html`<span class="badge free">Available</span>`;
  return html`<span class="badge locked" title="Lock expires ${f.lock.expiresAt}" data-expires="${f.lock.expiresAt}"
    >Being edited by ${f.lock.user.displayName} since ${time(f.lock.since, 'time')}</span>`;
}

function editButton(f: FileSummary): SafeHtml {
  if (!f.type) return html``;
  return html`<button class="button" data-action="edit" data-file-id="${f.id}" data-locked-by="${f.lock?.user.displayName ?? ''}">
    Edit in ${f.type.label}
  </button>`;
}

export function layout(opts: {
  title: string;
  user: AuthUser | null;
  page: string;
  fileId?: string;
  body: SafeHtml;
}): string {
  return html`<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${opts.title} · Office DAV demo</title>
  <link rel="stylesheet" href="/static/style.css">
</head>
<body data-page="${opts.page}" data-file-id="${opts.fileId ?? ''}">
  <header class="top">
    <a href="/" class="brand">Office DAV demo</a>
    <span class="spacer"></span>
    ${
      opts.user
        ? html`<span id="live" class="live" title="Live updates (SSE)">live</span>
            <span class="who">Signed in as <strong>${opts.user.displayName}</strong></span>
            <form method="post" action="/logout"><button class="link">Log out</button></form>`
        : ''
    }
  </header>
  <main>${opts.body}</main>
  <div id="toast" class="toast" role="status" hidden></div>
  <script src="/static/app.js" defer></script>
</body>
</html>`.value;
}

export function loginPage(users: readonly AuthUser[]): string {
  return layout({
    title: 'Sign in',
    user: null,
    page: 'login',
    body: html`<section class="card narrow">
      <h1>Sign in</h1>
      <p class="muted">Mock OIDC login. Pick a user; the real app uses its SSO here.</p>
      <form method="post" action="/login" class="user-picker">
        ${users.map((u) => html`<button class="button" name="userId" value="${u.id}">${u.displayName}</button>`)}
      </form>
    </section>`,
  });
}

export function fileRows(files: FileSummary[]): SafeHtml {
  if (files.length === 0) return html`<tr><td colspan="5" class="muted">No files yet. Upload one.</td></tr>`;
  return html`${files.map(
    (f) => html`<tr data-file-id="${f.id}">
      <td><a class="file-name" href="/files/${f.id}">${fileIcon(f)}<span dir="auto">${f.name}</span></a></td>
      <td>${lockBadge(f)}</td>
      <td class="num">${f.versionCount}</td>
      <td>${time(f.updatedAt)}${f.updatedBy ? html` <span class="muted">by ${f.updatedBy.displayName}</span>` : ''}</td>
      <td class="actions">${editButton(f)}</td>
    </tr>`,
  )}`;
}

export function filesPage(user: AuthUser, files: FileSummary[]): string {
  return layout({
    title: 'Files',
    user,
    page: 'list',
    body: html`<section class="card">
      <div class="card-head">
        <h1>Files</h1>
        <label class="button secondary">Upload a file
          <input type="file" accept=".docx,.xlsx,.pptx" data-upload-new hidden>
        </label>
      </div>
      <table class="files">
        <thead><tr><th>Name</th><th>Status</th><th class="num">Versions</th><th>Last saved</th><th></th></tr></thead>
        <tbody id="file-rows">${fileRows(files)}</tbody>
      </table>
      <p class="muted small">Edit opens the file in Office desktop over WebDAV. Each Ctrl+S shows up here without a reload.</p>
    </section>`,
  });
}

export function fileDetails(d: FileDetail): SafeHtml {
  const live = d.sessions.filter((s) => !s.revokedAt);
  return html`<section class="card">
    <div class="card-head">
      <div>
        <h1 class="file-title">${fileIcon(d)}<span dir="auto">${d.name}</span></h1>
        <p class="meta">
          ${lockBadge(d)} · v${d.currentVersionNo} · ${formatBytes(d.size)} · saved ${time(d.updatedAt)}
          ${d.updatedBy ? html`by ${d.updatedBy.displayName}` : ''}
        </p>
      </div>
      <div class="actions">
        ${editButton(d)}
        <label class="button secondary">Upload new version
          <input type="file" accept=".${d.type?.extension ?? 'docx'}" data-upload-version="${d.id}" hidden>
        </label>
      </div>
    </div>

    <h2>Versions <span class="muted small">(${d.versions.length})</span></h2>
    <table class="versions">
      <thead><tr><th>#</th><th>Saved</th><th>By</th><th class="num">Size</th><th>Source</th><th></th></tr></thead>
      <tbody>
        ${d.versions.map(
          (v) => html`<tr data-version-id="${v.id}" class="${v.isCurrent ? 'current' : ''}">
            <td>v${v.versionNo}</td>
            <td>${time(v.createdAt)}</td>
            <td>${v.createdBy.displayName}</td>
            <td class="num">${formatBytes(v.size)}</td>
            <td><span class="source ${v.source}">${v.source}</span></td>
            <td><a href="/files/${d.id}/versions/${v.id}/download">Download</a></td>
          </tr>`,
        )}
      </tbody>
    </table>

    <h2>Admin <span class="muted small">for testing the failure paths</span></h2>
    <div class="admin">
      <button class="button danger" data-action="force-unlock" data-file-id="${d.id}" ${d.lock ? '' : 'disabled'}>Force unlock</button>
      <table class="sessions">
        <thead><tr><th>DAV session</th><th>User</th><th>Opened</th><th>Last LOCK/PUT</th><th>Lock</th><th></th></tr></thead>
        <tbody>
          ${
            live.length === 0
              ? html`<tr><td colspan="6" class="muted">No live DAV sessions.</td></tr>`
              : live.map(
                  (s) => html`<tr>
                    <td><code>${s.id.slice(0, 8)}</code></td>
                    <td>${s.user.displayName}</td>
                    <td>${time(s.createdAt)}</td>
                    <td>${time(s.lastSeenAt)}</td>
                    <td>${s.holdsLock ? 'holds lock' : ''}</td>
                    <td><button class="button danger small" data-action="revoke" data-session-id="${s.id}">Revoke session</button></td>
                  </tr>`,
                )
          }
        </tbody>
      </table>
    </div>
  </section>`;
}

/** `preview` is trusted HTML from PreviewService. */
export function filePage(user: AuthUser, detail: FileDetail, preview: string): string {
  return layout({
    title: detail.name,
    user,
    page: 'file',
    fileId: detail.id,
    body: html`<p class="crumbs"><a href="/">← All files</a></p>
      <div id="file-details">${fileDetails(detail)}</div>
      <div class="grid">
        <section class="card">
          <h2>Preview</h2>
          <div id="preview" class="preview">${raw(preview)}</div>
        </section>
        <section class="card">
          <h2>Live activity</h2>
          <ol id="activity" class="activity"><li class="muted">Waiting for events…</li></ol>
        </section>
      </div>`,
  });
}

export function errorPage(user: AuthUser | null, status: number, message: string): string {
  return layout({
    title: `Error ${status}`,
    user,
    page: 'error',
    body: html`<section class="card narrow"><h1>${status}</h1><p>${message}</p><p><a href="/">Back to files</a></p></section>`,
  });
}
