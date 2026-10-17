import type { Readable } from 'node:stream';
import JSZip from 'jszip';
import mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import type { FileService } from '../domain/files.js';
import { officeTypeFor } from '../domain/office-types.js';
import type { Repo } from '../storage/repo.js';
import { formatBytes, html, raw, type SafeHtml } from './html.js';

const MAX_PREVIEW_BYTES = 20 * 1024 * 1024;
const MAX_ROWS = 100;
const CACHE_SIZE = 50;

async function readAll(stream: Readable): Promise<Buffer> {
  const chunks: Buffer[] = [];
  for await (const chunk of stream) chunks.push(chunk as Buffer);
  return Buffer.concat(chunks);
}

/**
 * mammoth and SheetJS escape text, but links and images come from the
 * document. Keep only harmless URL schemes and drop event handlers.
 */
export function sanitizeGeneratedHtml(input: string): string {
  return input
    .replace(/\s+on[a-z]+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, '')
    .replace(/\s+href\s*=\s*"([^"]*)"/gi, (m, url: string) => (/^(https?:|mailto:|#)/i.test(url.trim()) ? m : ''))
    .replace(/\s+src\s*=\s*"([^"]*)"/gi, (m, url: string) => (/^data:image\/(png|jpe?g|gif|bmp);/i.test(url.trim()) ? m : ''))
    .replace(/<\/?(script|style|iframe|object|embed|form|link|meta)\b[^>]*>/gi, '');
}

async function docxPreview(buffer: Buffer): Promise<SafeHtml> {
  const result = await mammoth.convertToHtml({ buffer });
  return html`<div class="doc" dir="auto">${raw(sanitizeGeneratedHtml(result.value))}</div>`;
}

function xlsxPreview(buffer: Buffer): SafeHtml {
  const workbook = XLSX.read(buffer, { type: 'buffer', sheetRows: MAX_ROWS + 1, cellHTML: false });
  const name = workbook.SheetNames[0];
  const sheet = name ? workbook.Sheets[name] : undefined;
  if (!name || !sheet) return html`<p class="muted">The workbook has no sheets.</p>`;
  const table = XLSX.utils.sheet_to_html(sheet, { header: '', footer: '' });
  return html`<p class="muted small">Sheet “${name}”, first ${MAX_ROWS} rows</p>
    <div class="sheet" dir="auto">${raw(sanitizeGeneratedHtml(table))}</div>`;
}

async function pptxMetadata(buffer: Buffer): Promise<Array<[string, string]>> {
  const zip = await JSZip.loadAsync(buffer);
  const app = (await zip.file('docProps/app.xml')?.async('string')) ?? '';
  const core = (await zip.file('docProps/core.xml')?.async('string')) ?? '';
  const pick = (xml: string, tag: string) => new RegExp(`<${tag}[^>]*>([^<]*)</${tag}>`).exec(xml)?.[1];
  const rows: Array<[string, string]> = [];
  const slides = pick(app, 'Slides') ?? String(Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)).length);
  rows.push(['Slides', slides]);
  const title = pick(core, 'dc:title');
  if (title) rows.push(['Title', title]);
  const author = pick(core, 'cp:lastModifiedBy') ?? pick(core, 'dc:creator');
  if (author) rows.push(['Last modified by (in Office)', author]);
  const application = pick(app, 'Application');
  if (application) rows.push(['Application', application]);
  return rows;
}

/** Renders the current version of a file as an HTML fragment. Cached per immutable version. */
export class PreviewService {
  private readonly cache = new Map<string, string>();

  constructor(
    private readonly files: FileService,
    private readonly repo: Repo,
  ) {}

  async render(fileId: string): Promise<string> {
    const { file, version, stream } = await this.files.openVersion(fileId);
    const cached = this.cache.get(version.id);
    if (cached) {
      stream.destroy();
      return cached;
    }

    const type = officeTypeFor(file.name);
    const author = this.repo.getUser(version.created_by)?.display_name ?? version.created_by;
    const header = html`<p class="preview-meta">Version ${version.version_no} · ${formatBytes(version.size)} · saved
      <time datetime="${new Date(version.created_at).toISOString()}" data-fmt="datetime"></time> by ${author}</p>`;

    let body: SafeHtml;
    if (version.size > MAX_PREVIEW_BYTES && type?.extension !== 'pptx') {
      stream.destroy();
      body = html`<p class="muted">Too large to preview (${formatBytes(version.size)}).</p>`;
    } else {
      try {
        const buffer = await readAll(stream);
        if (type?.extension === 'docx') body = await docxPreview(buffer);
        else if (type?.extension === 'xlsx') body = xlsxPreview(buffer);
        else if (type?.extension === 'pptx') {
          const rows = await pptxMetadata(buffer);
          body = html`<table class="meta-table">${rows.map(([k, v]) => html`<tr><th>${k}</th><td dir="auto">${v}</td></tr>`)}</table>
            <p class="muted small">PowerPoint files are not rendered; this is the file's metadata.</p>`;
        } else body = html`<p class="muted">No preview for this file type.</p>`;
      } catch {
        body = html`<p class="muted">Preview unavailable: this version could not be parsed.</p>`;
      }
    }

    const fragment = html`${header}${body}`.value;
    this.cache.set(version.id, fragment);
    while (this.cache.size > CACHE_SIZE) this.cache.delete(this.cache.keys().next().value!);
    return fragment;
  }
}

