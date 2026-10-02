import { ZipReader } from './zipReader';

export interface WordMetadata {
  title?: string;
  creator?: string;
  lastModifiedBy?: string;
  created?: string;
  modified?: string;
  revision?: string;
  words?: number;
  characters?: number;
  pages?: number;
  paragraphs?: number;
  fileSizeBytes?: number;
  isLegacyDoc?: boolean;
}

export interface WordHeading {
  id: string;
  text: string;
  level: number;
}

export interface WordRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  highlight?: string;
  fontSizePt?: number;
  fontFamily?: string;
  superscript?: boolean;
  subscript?: boolean;
  href?: string;
}

export interface WordParagraphElement {
  id: string;
  type: 'paragraph';
  headingLevel?: number; // 1-6
  isTitle?: boolean;
  isSubtitle?: boolean;
  alignment: 'left' | 'center' | 'right' | 'justify';
  isRtl: boolean;
  isList?: boolean;
  listType?: 'bullet' | 'number';
  listLevel?: number;
  indentPt?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: number;
  runs: WordRun[];
}

export interface WordTableCell {
  id: string;
  colspan?: number;
  rowspan?: number;
  backgroundColor?: string;
  borderColor?: string;
  widthPercent?: number;
  elements: (WordParagraphElement | WordTableElement | WordImageElement)[];
}

export interface WordTableRow {
  id: string;
  cells: WordTableCell[];
  isHeader?: boolean;
}

export interface WordTableElement {
  id: string;
  type: 'table';
  rows: WordTableRow[];
  alignment: 'left' | 'center' | 'right';
  widthPercent?: number;
}

export interface WordImageElement {
  id: string;
  type: 'image';
  src: string;
  alt?: string;
  widthPx?: number;
  heightPx?: number;
  alignment: 'left' | 'center' | 'right';
}

export interface WordPageBreakElement {
  id: string;
  type: 'pageBreak';
}

export type WordDocElement =
  | WordParagraphElement
  | WordTableElement
  | WordImageElement
  | WordPageBreakElement;

export interface WordPage {
  pageNumber: number;
  elements: WordDocElement[];
}

export interface WordParsedDocument {
  metadata: WordMetadata;
  headings: WordHeading[];
  pages: WordPage[];
  allElements: WordDocElement[];
  plainText: string;
  hasImages: boolean;
}

/**
 * Checks whether text contains Arabic / Hebrew RTL characters.
 */
export function isRtlText(text: string): boolean {
  return /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF\u0590-\u05FF]/.test(
    text
  );
}

/**
 * Parses a Word document ArrayBuffer (.docx or legacy .doc).
 */
export async function parseWordDocument(buffer: ArrayBuffer): Promise<WordParsedDocument> {
  const bytes = new Uint8Array(buffer);

  // Check magic bytes
  // 1. DOCX (ZIP archive PK\x03\x04: 0x50 0x4B 0x03 0x04)
  if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04) {
    return parseDocxArchive(buffer);
  }

  // 2. Legacy binary .doc (OLE2 Compound Document: 0xD0 0xCF 0x11 0xE0 0xA1 0xB1 0x1A 0xE1)
  if (
    bytes.length >= 8 &&
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0 &&
    bytes[4] === 0xa1 &&
    bytes[5] === 0xb1 &&
    bytes[6] === 0x1a &&
    bytes[7] === 0xe1
  ) {
    return parseLegacyDoc(bytes);
  }

  // 3. HTML disguised as .doc (common in many export engines)
  const previewStr = new TextDecoder('utf-8', { fatal: false }).decode(bytes.slice(0, 1024)).trim().toLowerCase();
  if (previewStr.includes('<html') || previewStr.includes('<!doctype html') || previewStr.includes('<body')) {
    return parseHtmlDisguisedDoc(bytes);
  }

  // 4. RTF format (starts with {\rtf)
  if (previewStr.startsWith('{\\rtf')) {
    return parseRtfDoc(bytes);
  }

  // Try treating as ZIP anyway (some have slight header offsets)
  try {
    return await parseDocxArchive(buffer);
  } catch {
    // Fall back to raw text extraction
    return parseRawFallback(bytes);
  }
}

/**
 * Parse modern .docx OpenXML package
 */
async function parseDocxArchive(buffer: ArrayBuffer): Promise<WordParsedDocument> {
  const zip = new ZipReader(buffer);
  await zip.init();

  // 1. Read document relationships (_rels/document.xml.rels)
  const relsMap = new Map<string, { target: string; type: string }>();
  const relsXmlStr =
    (await zip.getText('word/_rels/document.xml.rels')) ||
    (await zip.getText('word/_rels/document.xml.rels', 'utf-8'));

  if (relsXmlStr) {
    const parser = new DOMParser();
    const relsDoc = parser.parseFromString(relsXmlStr, 'application/xml');
    const relationships = relsDoc.getElementsByTagName('Relationship');
    for (let i = 0; i < relationships.length; i++) {
      const el = relationships[i];
      const id = el.getAttribute('Id') || '';
      const target = el.getAttribute('Target') || '';
      const type = el.getAttribute('Type') || '';
      if (id && target) {
        relsMap.set(id, { target, type });
      }
    }
  }

  // 2. Read images into Blob URLs
  const imageBlobUrlMap = new Map<string, string>();
  for (const [rId, rel] of relsMap.entries()) {
    if (rel.type.includes('image') || /\.(png|jpe?g|gif|svg|webp|bmp)$/i.test(rel.target)) {
      let targetPath = rel.target;
      if (!targetPath.startsWith('word/')) {
        targetPath = 'word/' + targetPath.replace(/^\//, '');
      }
      const blobUrl = await zip.getBlobUrl(targetPath);
      if (blobUrl) {
        imageBlobUrlMap.set(rId, blobUrl);
      }
    }
  }

  // 3. Read metadata from docProps/core.xml & docProps/app.xml
  const metadata: WordMetadata = {
    fileSizeBytes: buffer.byteLength,
  };

  const coreXmlStr = await zip.getText('docProps/core.xml');
  if (coreXmlStr) {
    const parser = new DOMParser();
    const coreDoc = parser.parseFromString(coreXmlStr, 'application/xml');
    metadata.title = coreDoc.querySelector('title')?.textContent || undefined;
    metadata.creator = coreDoc.querySelector('creator')?.textContent || undefined;
    metadata.lastModifiedBy = coreDoc.querySelector('lastModifiedBy')?.textContent || undefined;
    metadata.created = coreDoc.querySelector('created')?.textContent || undefined;
    metadata.modified = coreDoc.querySelector('modified')?.textContent || undefined;
    metadata.revision = coreDoc.querySelector('revision')?.textContent || undefined;
  }

  const appXmlStr = await zip.getText('docProps/app.xml');
  if (appXmlStr) {
    const parser = new DOMParser();
    const appDoc = parser.parseFromString(appXmlStr, 'application/xml');
    const words = parseInt(appDoc.querySelector('Words')?.textContent || '', 10);
    const chars = parseInt(appDoc.querySelector('Characters')?.textContent || '', 10);
    const pages = parseInt(appDoc.querySelector('Pages')?.textContent || '', 10);
    const paragraphs = parseInt(appDoc.querySelector('Paragraphs')?.textContent || '', 10);

    if (!isNaN(words)) metadata.words = words;
    if (!isNaN(chars)) metadata.characters = chars;
    if (!isNaN(pages)) metadata.pages = pages;
    if (!isNaN(paragraphs)) metadata.paragraphs = paragraphs;
  }

  // 4. Read main document: word/document.xml
  const docXmlStr = await zip.getText('word/document.xml');
  if (!docXmlStr) {
    throw new Error('Invalid Word document: missing word/document.xml');
  }

  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(docXmlStr, 'application/xml');

  // Check for parse error
  const parseError = xmlDoc.querySelector('parsererror');
  if (parseError) {
    throw new Error('Error parsing Word XML document: ' + parseError.textContent);
  }

  const bodyEl = xmlDoc.getElementsByTagName('w:body')[0] || xmlDoc.querySelector('body');
  if (!bodyEl) {
    throw new Error('Document body missing in word/document.xml');
  }

  const headings: WordHeading[] = [];
  const allElements: WordDocElement[] = [];
  let elementCounter = 0;
  let hasImages = false;

  const childNodes = Array.from(bodyEl.childNodes);

  for (const node of childNodes) {
    if (node.nodeType !== Node.ELEMENT_NODE) continue;
    const el = node as Element;
    const nodeName = el.localName || el.nodeName.replace(/^.*:/, '');

    if (nodeName === 'p') {
      const pElem = parseParagraph(el, relsMap, imageBlobUrlMap, `p-${++elementCounter}`);
      if (pElem) {
        allElements.push(pElem);
        if (pElem.type === 'image') {
          hasImages = true;
        } else if (pElem.type === 'paragraph') {
          // Check if heading for outline
          if (pElem.headingLevel && pElem.runs.length > 0) {
            const headingText = pElem.runs.map((r) => r.text).join('').trim();
            if (headingText) {
              headings.push({
                id: pElem.id,
                text: headingText,
                level: pElem.headingLevel,
              });
            }
          }
        }
      }
    } else if (nodeName === 'tbl') {
      const tblElem = parseTable(el, relsMap, imageBlobUrlMap, `tbl-${++elementCounter}`);
      if (tblElem) {
        allElements.push(tblElem);
      }
    }
  }

  // Calculate plain text and words if missing
  const plainText = allElements
    .map((el) => {
      if (el.type === 'paragraph') {
        return el.runs.map((r) => r.text).join('');
      } else if (el.type === 'table') {
        return el.rows
          .map((row) =>
            row.cells
              .map((c) =>
                c.elements
                  .map((sub) => (sub.type === 'paragraph' ? sub.runs.map((r) => r.text).join('') : ''))
                  .join(' ')
              )
              .join('\t')
          )
          .join('\n');
      }
      return '';
    })
    .join('\n');

  if (!metadata.characters) {
    metadata.characters = plainText.length;
  }
  if (!metadata.words) {
    const wordMatches = plainText.match(/\S+/g);
    metadata.words = wordMatches ? wordMatches.length : 0;
  }
  if (!metadata.paragraphs) {
    metadata.paragraphs = allElements.filter((e) => e.type === 'paragraph').length;
  }

  // Split into simulated pages (based on explicit page breaks or ~500 words / element density)
  const pages = paginateElements(allElements);

  if (!metadata.pages || metadata.pages < 1) {
    metadata.pages = Math.max(1, pages.length);
  }

  return {
    metadata,
    headings,
    pages,
    allElements,
    plainText,
    hasImages: hasImages || imageBlobUrlMap.size > 0,
  };
}

/**
 * Parse `<w:p>` paragraph node
 */
function parseParagraph(
  pNode: Element,
  relsMap: Map<string, { target: string; type: string }>,
  imageBlobUrlMap: Map<string, string>,
  id: string
): WordParagraphElement | WordImageElement | WordPageBreakElement | null {
  // Check for standalone image inside drawing
  const drawing = pNode.querySelector('drawing, w\\:drawing') || pNode.getElementsByTagName('w:drawing')[0];
  if (drawing) {
    const imgElem = extractImageFromDrawing(drawing, relsMap, imageBlobUrlMap, id);
    if (imgElem) {
      return imgElem;
    }
  }

  // Check for explicit page break in paragraph
  const pageBr = pNode.querySelector('br[w\\:type="page"], w\\:br[w\\:type="page"]');
  if (pageBr) {
    return { id, type: 'pageBreak' };
  }

  const pPr = pNode.querySelector('pPr, w\\:pPr');
  let alignment: 'left' | 'center' | 'right' | 'justify' = 'left';
  let headingLevel: number | undefined = undefined;
  let isTitle = false;
  let isSubtitle = false;
  let isList = false;
  let listType: 'bullet' | 'number' = 'bullet';
  let listLevel = 0;
  let indentPt: number | undefined = undefined;
  let spacingBeforePt: number | undefined = undefined;
  let spacingAfterPt: number | undefined = undefined;
  let lineSpacing: number | undefined = undefined;
  let isRtl = false;

  if (pPr) {
    // Alignment
    const jc = pPr.querySelector('jc, w\\:jc');
    if (jc) {
      const val = (jc.getAttribute('w:val') || jc.getAttribute('val') || '').toLowerCase();
      if (val === 'center') alignment = 'center';
      else if (val === 'right') alignment = 'right';
      else if (val === 'both' || val === 'justify') alignment = 'justify';
      else if (val === 'left') alignment = 'left';
    }

    // Bidirectional / RTL
    if (pPr.querySelector('bidi, w\\:bidi')) {
      isRtl = true;
    }

    // Paragraph Style (Headings, Title)
    const pStyle = pPr.querySelector('pStyle, w\\:pStyle');
    if (pStyle) {
      const val = (pStyle.getAttribute('w:val') || pStyle.getAttribute('val') || '').toLowerCase();
      if (val.includes('heading1') || val === 'heading 1' || val === 'h1' || val === '1') headingLevel = 1;
      else if (val.includes('heading2') || val === 'heading 2' || val === 'h2' || val === '2') headingLevel = 2;
      else if (val.includes('heading3') || val === 'heading 3' || val === 'h3' || val === '3') headingLevel = 3;
      else if (val.includes('heading4') || val === 'heading 4' || val === 'h4' || val === '4') headingLevel = 4;
      else if (val.includes('heading5') || val === 'heading 5' || val === 'h5' || val === '5') headingLevel = 5;
      else if (val.includes('heading6') || val === 'heading 6' || val === 'h6' || val === '6') headingLevel = 6;
      else if (val.includes('title')) isTitle = true;
      else if (val.includes('subtitle')) isSubtitle = true;
    }

    // Numbering / Bullets
    const numPr = pPr.querySelector('numPr, w\\:numPr');
    if (numPr) {
      isList = true;
      const ilvl = numPr.querySelector('ilvl, w\\:ilvl');
      if (ilvl) {
        const lvlVal = parseInt(ilvl.getAttribute('w:val') || ilvl.getAttribute('val') || '0', 10);
        if (!isNaN(lvlVal)) listLevel = lvlVal;
      }
      const numId = numPr.querySelector('numId, w\\:numId');
      if (numId) {
        const idVal = numId.getAttribute('w:val') || numId.getAttribute('val');
        // If numId > 0, usually numbered or bullet list
        if (idVal && idVal !== '0') {
          listType = parseInt(idVal, 10) % 2 === 0 ? 'number' : 'bullet';
        }
      }
    }

    // Indentation
    const ind = pPr.querySelector('ind, w\\:ind');
    if (ind) {
      const leftDxa = parseInt(ind.getAttribute('w:left') || ind.getAttribute('left') || '0', 10);
      if (!isNaN(leftDxa) && leftDxa > 0) {
        indentPt = leftDxa / 20; // 20 dxa = 1 pt
      }
    }

    // Spacing
    const spacing = pPr.querySelector('spacing, w\\:spacing');
    if (spacing) {
      const beforeDxa = parseInt(spacing.getAttribute('w:before') || spacing.getAttribute('before') || '0', 10);
      const afterDxa = parseInt(spacing.getAttribute('w:after') || spacing.getAttribute('after') || '0', 10);
      const lineDxa = parseInt(spacing.getAttribute('w:line') || spacing.getAttribute('line') || '0', 10);
      if (!isNaN(beforeDxa) && beforeDxa > 0) spacingBeforePt = beforeDxa / 20;
      if (!isNaN(afterDxa) && afterDxa > 0) spacingAfterPt = afterDxa / 20;
      if (!isNaN(lineDxa) && lineDxa > 0) lineSpacing = lineDxa / 240; // 240 = single line
    }
  }

  // Parse runs
  const runs: WordRun[] = [];
  const children = Array.from(pNode.childNodes);

  for (const child of children) {
    if (child.nodeType !== Node.ELEMENT_NODE) continue;
    const childEl = child as Element;
    const childName = childEl.localName || childEl.nodeName.replace(/^.*:/, '');

    if (childName === 'r') {
      const run = parseRun(childEl);
      if (run) runs.push(run);
    } else if (childName === 'hyperlink') {
      const rId = childEl.getAttribute('r:id') || childEl.getAttribute('id') || '';
      const href = relsMap.get(rId)?.target;
      const subRuns = childEl.querySelectorAll('r, w\\:r');
      for (let i = 0; i < subRuns.length; i++) {
        const r = parseRun(subRuns[i]);
        if (r) {
          r.href = href;
          runs.push(r);
        }
      }
    }
  }

  // Determine RTL if any run contains Arabic text
  if (!isRtl) {
    const fullText = runs.map((r) => r.text).join('');
    if (isRtlText(fullText)) {
      isRtl = true;
      if (alignment === 'left') {
        alignment = 'right';
      }
    }
  }

  return {
    id,
    type: 'paragraph',
    headingLevel,
    isTitle,
    isSubtitle,
    alignment,
    isRtl,
    isList,
    listType,
    listLevel,
    indentPt,
    spacingBeforePt,
    spacingAfterPt,
    lineSpacing,
    runs,
  };
}

/**
 * Parse `<w:r>` text run
 */
function parseRun(rNode: Element): WordRun | null {
  const rPr = rNode.querySelector('rPr, w\\:rPr');
  let bold = false;
  let italic = false;
  let underline = false;
  let strike = false;
  let color: string | undefined = undefined;
  let highlight: string | undefined = undefined;
  let fontSizePt: number | undefined = undefined;
  let fontFamily: string | undefined = undefined;
  let superscript = false;
  let subscript = false;

  if (rPr) {
    bold = Boolean(rPr.querySelector('b, w\\:b'));
    italic = Boolean(rPr.querySelector('i, w\\:i'));
    underline = Boolean(rPr.querySelector('u, w\\:u'));
    strike = Boolean(rPr.querySelector('strike, w\\:strike'));

    const colorEl = rPr.querySelector('color, w\\:color');
    if (colorEl) {
      const hex = colorEl.getAttribute('w:val') || colorEl.getAttribute('val');
      if (hex && hex !== 'auto') {
        color = `#${hex}`;
      }
    }

    const hlEl = rPr.querySelector('highlight, w\\:highlight');
    if (hlEl) {
      highlight = hlEl.getAttribute('w:val') || hlEl.getAttribute('val') || undefined;
    }

    const szEl = rPr.querySelector('sz, w\\:sz');
    if (szEl) {
      const halfPts = parseInt(szEl.getAttribute('w:val') || szEl.getAttribute('val') || '0', 10);
      if (!isNaN(halfPts) && halfPts > 0) {
        fontSizePt = halfPts / 2; // 2 half-points = 1 pt
      }
    }

    const rFonts = rPr.querySelector('rFonts, w\\:rFonts');
    if (rFonts) {
      fontFamily =
        rFonts.getAttribute('w:ascii') ||
        rFonts.getAttribute('w:cs') ||
        rFonts.getAttribute('ascii') ||
        undefined;
    }

    const vertAlign = rPr.querySelector('vertAlign, w\\:vertAlign');
    if (vertAlign) {
      const val = vertAlign.getAttribute('w:val') || vertAlign.getAttribute('val');
      if (val === 'superscript') superscript = true;
      else if (val === 'subscript') subscript = true;
    }
  }

  // Extract text and breaks
  let text = '';
  const textChildren = Array.from(rNode.childNodes);

  for (const child of textChildren) {
    if (child.nodeType !== Node.ELEMENT_NODE) continue;
    const el = child as Element;
    const name = el.localName || el.nodeName.replace(/^.*:/, '');

    if (name === 't') {
      text += el.textContent || '';
    } else if (name === 'br' || name === 'cr') {
      text += '\n';
    } else if (name === 'tab') {
      text += '    ';
    }
  }

  if (!text) return null;

  return {
    text,
    bold,
    italic,
    underline,
    strike,
    color,
    highlight,
    fontSizePt,
    fontFamily,
    superscript,
    subscript,
  };
}

/**
 * Extract image element from `<w:drawing>`
 */
function extractImageFromDrawing(
  drawingNode: Element,
  relsMap: Map<string, { target: string; type: string }>,
  imageBlobUrlMap: Map<string, string>,
  id: string
): WordImageElement | null {
  const blip = drawingNode.querySelector('blip, a\\:blip');
  if (!blip) return null;

  const embedId =
    blip.getAttribute('r:embed') ||
    blip.getAttribute('embed') ||
    blip.getAttribute('r:link') ||
    '';

  const blobUrl = imageBlobUrlMap.get(embedId) || relsMap.get(embedId)?.target;
  if (!blobUrl) return null;

  // Read extent (cx, cy in EMUs)
  // 1 EMU = 1/914400 inch = 96/914400 px
  let widthPx: number | undefined = undefined;
  let heightPx: number | undefined = undefined;

  const ext = drawingNode.querySelector('extent, wp\\:extent, xfrm > ext, a\\:xfrm > a\\:ext');
  if (ext) {
    const cx = parseInt(ext.getAttribute('cx') || '0', 10);
    const cy = parseInt(ext.getAttribute('cy') || '0', 10);
    if (!isNaN(cx) && cx > 0) widthPx = Math.round((cx * 96) / 914400);
    if (!isNaN(cy) && cy > 0) heightPx = Math.round((cy * 96) / 914400);
  }

  const docPr = drawingNode.querySelector('docPr, wp\\:docPr');
  const alt = docPr?.getAttribute('name') || docPr?.getAttribute('descr') || 'Word document image';

  return {
    id,
    type: 'image',
    src: blobUrl,
    alt,
    widthPx: widthPx ? Math.min(widthPx, 750) : undefined,
    heightPx: heightPx ? Math.min(heightPx, 1000) : undefined,
    alignment: 'center',
  };
}

/**
 * Parse `<w:tbl>` table node
 */
function parseTable(
  tblNode: Element,
  relsMap: Map<string, { target: string; type: string }>,
  imageBlobUrlMap: Map<string, string>,
  id: string
): WordTableElement | null {
  let alignment: 'left' | 'center' | 'right' = 'left';
  const jc = tblNode.querySelector('tblPr > jc, w\\:tblPr > w\\:jc');
  if (jc) {
    const val = (jc.getAttribute('w:val') || jc.getAttribute('val') || '').toLowerCase();
    if (val === 'center') alignment = 'center';
    else if (val === 'right') alignment = 'right';
  }

  const rows: WordTableRow[] = [];
  const trList = tblNode.querySelectorAll('tr, w\\:tr');
  let rowCounter = 0;

  for (let r = 0; r < trList.length; r++) {
    const tr = trList[r];
    const isHeader = r === 0 || Boolean(tr.querySelector('tblHeader, w\\:tblHeader'));
    const cells: WordTableCell[] = [];
    const tcList = tr.querySelectorAll('tc, w\\:tc');
    let cellCounter = 0;

    for (let c = 0; c < tcList.length; c++) {
      const tc = tcList[c];
      const tcPr = tc.querySelector('tcPr, w\\:tcPr');
      let colspan = 1;
      let backgroundColor: string | undefined = undefined;

      if (tcPr) {
        const gridSpan = tcPr.querySelector('gridSpan, w\\:gridSpan');
        if (gridSpan) {
          const spanVal = parseInt(gridSpan.getAttribute('w:val') || gridSpan.getAttribute('val') || '1', 10);
          if (!isNaN(spanVal) && spanVal > 1) colspan = spanVal;
        }

        const shd = tcPr.querySelector('shd, w\\:shd');
        if (shd) {
          const fill = shd.getAttribute('w:fill') || shd.getAttribute('fill');
          if (fill && fill !== 'auto' && fill !== 'none') {
            backgroundColor = `#${fill}`;
          }
        }
      }

      // Parse nested paragraphs or elements in cell
      const cellElements: (WordParagraphElement | WordTableElement | WordImageElement)[] = [];
      const cellChildren = Array.from(tc.childNodes);
      let subCounter = 0;

      for (const node of cellChildren) {
        if (node.nodeType !== Node.ELEMENT_NODE) continue;
        const subEl = node as Element;
        const name = subEl.localName || subEl.nodeName.replace(/^.*:/, '');

        if (name === 'p') {
          const p = parseParagraph(subEl, relsMap, imageBlobUrlMap, `${id}-c${cellCounter}-p${++subCounter}`);
          if (p && p.type === 'paragraph') cellElements.push(p);
        }
      }

      cells.push({
        id: `${id}-r${rowCounter}-c${++cellCounter}`,
        colspan,
        backgroundColor,
        elements: cellElements,
      });
    }

    rows.push({
      id: `${id}-r${++rowCounter}`,
      isHeader,
      cells,
    });
  }

  return {
    id,
    type: 'table',
    alignment,
    rows,
  };
}

/**
 * Groups elements into discrete pages for the "Paged Mode" view.
 */
function paginateElements(elements: WordDocElement[]): WordPage[] {
  const pages: WordPage[] = [];
  let currentPageElements: WordDocElement[] = [];
  let pageNum = 1;
  let wordCountInPage = 0;

  const PAGE_WORD_LIMIT = 450; // Standard Word page density

  for (const el of elements) {
    if (el.type === 'pageBreak') {
      if (currentPageElements.length > 0) {
        pages.push({ pageNumber: pageNum++, elements: currentPageElements });
        currentPageElements = [];
        wordCountInPage = 0;
      }
      continue;
    }

    let elWords = 0;
    if (el.type === 'paragraph') {
      const txt = el.runs.map((r) => r.text).join(' ');
      elWords = (txt.match(/\S+/g) || []).length;
    } else if (el.type === 'table') {
      elWords = 100; // estimated table weight
    } else if (el.type === 'image') {
      elWords = 80;
    }

    if (wordCountInPage > 0 && wordCountInPage + elWords > PAGE_WORD_LIMIT) {
      pages.push({ pageNumber: pageNum++, elements: currentPageElements });
      currentPageElements = [el];
      wordCountInPage = elWords;
    } else {
      currentPageElements.push(el);
      wordCountInPage += elWords;
    }
  }

  if (currentPageElements.length > 0) {
    pages.push({ pageNumber: pageNum, elements: currentPageElements });
  }

  if (pages.length === 0) {
    pages.push({ pageNumber: 1, elements: [] });
  }

  return pages;
}

/**
 * Fallback parser for legacy .doc (Compound File Binary Format / OLE2)
 */
function parseLegacyDoc(bytes: Uint8Array): WordParsedDocument {
  // Extract printable Unicode UTF-16LE and ASCII text strings
  const paragraphs: string[] = [];
  let currentWord = '';

  // Scan for UTF-16LE strings (common in WordDocument stream)
  for (let i = 0; i < bytes.length - 1; i += 2) {
    const code = bytes[i] | (bytes[i + 1] << 8);
    // Printable characters + Arabic range + whitespace
    if (
      (code >= 32 && code <= 126) ||
      (code >= 0x0600 && code <= 0x06ff) ||
      code === 10 ||
      code === 13 ||
      code === 9
    ) {
      if (code === 10 || code === 13) {
        if (currentWord.trim().length > 0) {
          paragraphs.push(currentWord.trim());
          currentWord = '';
        }
      } else {
        currentWord += String.fromCharCode(code);
      }
    } else {
      if (currentWord.trim().length > 3) {
        paragraphs.push(currentWord.trim());
      }
      currentWord = '';
    }
  }

  if (currentWord.trim().length > 0) {
    paragraphs.push(currentWord.trim());
  }

  // Filter out binary garbage strings
  const cleanParagraphs = paragraphs.filter((p) => p.length >= 2 && !/^[\x00-\x1F\x7F]+$/.test(p));

  const allElements: WordDocElement[] = cleanParagraphs.map((text, idx) => {
    const isRtl = isRtlText(text);
    return {
      id: `doc-p-${idx + 1}`,
      type: 'paragraph',
      alignment: isRtl ? 'right' : 'left',
      isRtl,
      runs: [{ text }],
    };
  });

  const plainText = cleanParagraphs.join('\n\n');
  const words = (plainText.match(/\S+/g) || []).length;

  return {
    metadata: {
      fileSizeBytes: bytes.length,
      words,
      characters: plainText.length,
      paragraphs: cleanParagraphs.length,
      isLegacyDoc: true,
    },
    headings: [],
    pages: paginateElements(allElements),
    allElements,
    plainText,
    hasImages: false,
  };
}

/**
 * Fallback parser for HTML disguised as Word doc
 */
function parseHtmlDisguisedDoc(bytes: Uint8Array): WordParsedDocument {
  const html = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');

  const allElements: WordDocElement[] = [];
  const headings: WordHeading[] = [];
  let idCounter = 0;

  const bodyChildren = Array.from(doc.body.children);
  for (const child of bodyChildren) {
    const tagName = child.tagName.toLowerCase();
    const text = child.textContent?.trim() || '';
    if (!text) continue;

    const isRtl = isRtlText(text);
    let headingLevel: number | undefined = undefined;
    if (tagName === 'h1') headingLevel = 1;
    else if (tagName === 'h2') headingLevel = 2;
    else if (tagName === 'h3') headingLevel = 3;

    const pElem: WordParagraphElement = {
      id: `html-el-${++idCounter}`,
      type: 'paragraph',
      headingLevel,
      alignment: isRtl ? 'right' : 'left',
      isRtl,
      runs: [{ text }],
    };

    allElements.push(pElem);
    if (headingLevel) {
      headings.push({ id: pElem.id, text, level: headingLevel });
    }
  }

  const plainText = allElements.map((e) => (e.type === 'paragraph' ? e.runs[0]?.text : '')).join('\n');

  return {
    metadata: {
      title: doc.title || undefined,
      fileSizeBytes: bytes.length,
      words: (plainText.match(/\S+/g) || []).length,
      characters: plainText.length,
      paragraphs: allElements.length,
    },
    headings,
    pages: paginateElements(allElements),
    allElements,
    plainText,
    hasImages: false,
  };
}

/**
 * Fallback parser for RTF format
 */
function parseRtfDoc(bytes: Uint8Array): WordParsedDocument {
  const rtf = new TextDecoder('latin1').decode(bytes);
  // Strip RTF control words
  const clean = rtf
    .replace(/\\par[d]?/g, '\n')
    .replace(/\\[a-z0-9\-]+/gi, ' ')
    .replace(/[{}]/g, '')
    .trim();

  const lines = clean.split('\n').filter((l) => l.trim().length > 0);
  const allElements: WordDocElement[] = lines.map((text, idx) => ({
    id: `rtf-${idx + 1}`,
    type: 'paragraph',
    alignment: isRtlText(text) ? 'right' : 'left',
    isRtl: isRtlText(text),
    runs: [{ text }],
  }));

  const plainText = lines.join('\n');
  return {
    metadata: {
      fileSizeBytes: bytes.length,
      words: (plainText.match(/\S+/g) || []).length,
      characters: plainText.length,
      paragraphs: lines.length,
    },
    headings: [],
    pages: paginateElements(allElements),
    allElements,
    plainText,
    hasImages: false,
  };
}

/**
 * Generic binary fallback when all else fails
 */
function parseRawFallback(bytes: Uint8Array): WordParsedDocument {
  return parseLegacyDoc(bytes);
}
