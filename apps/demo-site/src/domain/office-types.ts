import type { OfficeApp } from '@office-dav/gateway-contract';

export interface OfficeType {
  extension: 'docx' | 'xlsx' | 'pptx';
  mime: string;
  app: OfficeApp;
  label: 'Word' | 'Excel' | 'PowerPoint';
}

const TYPES: Record<string, OfficeType> = {
  docx: {
    extension: 'docx',
    mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    app: 'word',
    label: 'Word',
  },
  xlsx: {
    extension: 'xlsx',
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    app: 'excel',
    label: 'Excel',
  },
  pptx: {
    extension: 'pptx',
    mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    app: 'powerpoint',
    label: 'PowerPoint',
  },
};

export function officeTypeFor(fileName: string): OfficeType | null {
  const ext = fileName.toLowerCase().split('.').pop() ?? '';
  return TYPES[ext] ?? null;
}
