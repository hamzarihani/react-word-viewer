export { WordViewer, WordViewer as CustomWordViewer, default } from './WordViewer';
export type {
  WordViewerProps,
  WordViewerProps as CustomWordViewerProps,
  WordViewerDictionary,
} from './WordViewer';

export {
  parseWordDocument,
  isRtlText,
} from './word/docxParser';

export type {
  WordParsedDocument,
  WordMetadata,
  WordHeading,
  WordDocElement,
  WordParagraphElement,
  WordTableElement,
  WordTableCell,
  WordTableRow,
  WordImageElement,
  WordPageBreakElement,
  WordRun,
  WordPage,
} from './word/docxParser';

export { ZipReader } from './word/zipReader';
export type { ZipEntry } from './word/zipReader';
