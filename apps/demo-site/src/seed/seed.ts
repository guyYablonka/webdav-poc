import { Readable } from 'node:stream';
import { AlignmentType, Document, HeadingLevel, Packer, Paragraph, TextRun } from 'docx';
import ExcelJS from 'exceljs';
import PptxGenJSModule from 'pptxgenjs';
import type { FastifyBaseLogger } from 'fastify';
import type { FileService } from '../domain/files.js';
import type { Repo } from '../storage/repo.js';

// pptxgenjs loads as ESM at runtime but ships CommonJS-style types, so under
// NodeNext the class appears one level down. The cast reconciles the two.
const PptxGenJS = PptxGenJSModule as unknown as typeof PptxGenJSModule.default;

async function projectPlanDocx(): Promise<Buffer> {
  const doc = new Document({
    creator: 'Office DAV demo',
    title: 'Project Plan',
    sections: [
      {
        children: [
          new Paragraph({ text: 'Project Plan', heading: HeadingLevel.HEADING_1 }),
          new Paragraph({ text: 'Edit this document in Word and press Ctrl+S. The demo site shows each save as a new version.' }),
          new Paragraph({ text: 'Milestones', heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: 'Discovery – done', bullet: { level: 0 } }),
          new Paragraph({ text: 'WebDAV POC – in progress', bullet: { level: 0 } }),
          new Paragraph({ text: 'Production rollout – planned', bullet: { level: 0 } }),
        ],
      },
    ],
  });
  return Packer.toBuffer(doc);
}

async function hebrewDocx(): Promise<Buffer> {
  const rtl = (text: string, bold = false) =>
    new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, children: [new TextRun({ text, rightToLeft: true, bold })] });
  const doc = new Document({
    creator: 'Office DAV demo',
    title: 'תוכנית עבודה',
    sections: [
      {
        children: [
          rtl('תוכנית עבודה (טיוטה)', true),
          rtl('מסמך לבדיקת שמות קבצים בעברית, עם רווחים וסוגריים.'),
          rtl('ערכו את המסמך ב-Word ולחצו Ctrl+S.'),
        ],
      },
    ],
  });
  return Packer.toBuffer(doc);
}

async function budgetXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Office DAV demo';
  const sheet = workbook.addWorksheet('Budget');
  sheet.columns = [
    { header: 'Item', key: 'item', width: 28 },
    { header: 'Q1', key: 'q1', width: 12 },
    { header: 'Q2', key: 'q2', width: 12 },
    { header: 'Q3', key: 'q3', width: 12 },
    { header: 'Q4', key: 'q4', width: 12 },
  ];
  sheet.addRows([
    { item: 'Salaries', q1: 120000, q2: 125000, q3: 130000, q4: 130000 },
    { item: 'Cloud', q1: 8000, q2: 8500, q3: 9000, q4: 9500 },
    { item: 'Licenses', q1: 4000, q2: 4000, q3: 4000, q4: 4000 },
    { item: 'Travel', q1: 2500, q2: 1500, q3: 3000, q4: 2000 },
  ]);
  sheet.addRow({
    item: 'Total',
    q1: { formula: 'SUM(B2:B5)' },
    q2: { formula: 'SUM(C2:C5)' },
    q3: { formula: 'SUM(D2:D5)' },
    q4: { formula: 'SUM(E2:E5)' },
  });
  sheet.getRow(1).font = { bold: true };
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

async function reviewPptx(): Promise<Buffer> {
  const pptx = new PptxGenJS();
  pptx.title = 'Quarterly Review';
  pptx.author = 'Office DAV demo';
  const title = pptx.addSlide();
  title.addText('Quarterly Review', { x: 0.5, y: 1.5, w: 9, h: 1.2, fontSize: 40, bold: true });
  title.addText('Edit in PowerPoint, then press Ctrl+S', { x: 0.5, y: 2.8, w: 9, h: 0.6, fontSize: 20 });
  const agenda = pptx.addSlide();
  agenda.addText('Agenda', { x: 0.5, y: 0.4, w: 9, h: 0.8, fontSize: 32, bold: true });
  agenda.addText(
    [
      { text: 'Results', options: { bullet: true } },
      { text: 'Risks', options: { bullet: true } },
      { text: 'Next quarter', options: { bullet: true } },
    ],
    { x: 0.7, y: 1.4, w: 8, h: 2.5, fontSize: 22 },
  );
  return (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
}

/** On first start, creates one docx, xlsx and pptx, plus a docx with a Hebrew name. */
export async function seedIfEmpty(files: FileService, repo: Repo, log: FastifyBaseLogger): Promise<void> {
  if (repo.countFiles() > 0) return;
  const seeds: Array<[string, () => Promise<Buffer>]> = [
    ['Project Plan.docx', projectPlanDocx],
    ['Budget 2026.xlsx', budgetXlsx],
    ['Quarterly Review.pptx', reviewPptx],
    ['תוכנית עבודה (טיוטה).docx', hebrewDocx],
  ];
  for (const [name, make] of seeds) {
    await files.createFile({ name, body: Readable.from([await make()]), userId: 'alice', source: 'seed' });
  }
  log.info({ files: seeds.map(([name]) => name) }, 'seeded demo files');
}
