import JSZip from 'jszip';
import {
  R,
  W,
  WmlPackage,
  XAttribute,
  XDeclaration,
  XDocument,
  XElement,
  XName,
  XNamespace,
} from 'openxmlsdkts';
import {
  CONTENT_TYPES,
  DOCUMENT_RELS,
  PACKAGE_RELS,
  SEED_DOCUMENT,
  SEED_NUMBERING,
  SEED_SETTINGS,
  SEED_STYLES,
} from './docx-seed.js';
import { PagekitError } from './errors.js';
import { characterStyleOf, halfPoints, mergeStyles, normalizeStyleInput, paragraphStyleOf, stripIllegalXml, twips } from './units.js';
import type {
  Block,
  DocxTemplate,
  DocumentTree,
  HeadingLevel,
  InlineNode,
  ListBlock,
  ListItem,
  ParagraphBlock,
  ResolvedStyle,
  TableBlock,
  WriteOptions,
  WriteResult,
} from './types.js';

const W_NAME = W.namespace.getName('name');

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const HYPERLINK_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const CONTENT_WIDTH = 9360;

const DEFAULT_NORMAL = normalizeStyleInput({
  fontFamily: 'Calibri',
  fontSize: 11,
  color: '#222222',
  spaceAfter: 8,
});

const DEFAULT_HEADINGS: Record<HeadingLevel, ResolvedStyle> = {
  1: normalizeStyleInput({ fontSize: 24, bold: true, color: '#1F4E79', spaceBefore: 12, spaceAfter: 6 }),
  2: normalizeStyleInput({ fontSize: 18, bold: true, color: '#1F4E79', spaceBefore: 10, spaceAfter: 4 }),
  3: normalizeStyleInput({ fontSize: 14, bold: true, color: '#1F4E79', spaceBefore: 8, spaceAfter: 4 }),
  4: normalizeStyleInput({ fontSize: 12, bold: true, color: '#1F4E79', spaceBefore: 8, spaceAfter: 2 }),
  5: normalizeStyleInput({ fontSize: 11, bold: true, italic: true, color: '#1F4E79', spaceBefore: 6, spaceAfter: 2 }),
  6: normalizeStyleInput({ fontSize: 11, italic: true, color: '#1F4E79', spaceBefore: 6, spaceAfter: 2 }),
};

const BULLETS = [
  { text: '\uF0B7', font: 'Symbol' },
  { text: 'o', font: 'Courier New' },
  { text: '\uF0A7', font: 'Wingdings' },
];

type TemplateLayers = {
  normal: ResolvedStyle;
  headings: Record<HeadingLevel, ResolvedStyle>;
};

type ListUse = {
  id: number;
  ordered: boolean;
  ilvl: number;
  start: number;
};

type WriteContext = {
  layers: TemplateLayers;
  defaults: ResolvedStyle;
  numbering: ListUse[];
  links: Map<string, string>;
  nextLink: number;
};

/**
 * Write a document tree to a .docx Blob.
 * Phase 1 only supports the `docx` format. The returned `skipped` list is
 * everything the tree could not represent, including items recorded by `parse`.
 */
export async function write(tree: DocumentTree, format: string, options?: WriteOptions): Promise<WriteResult> {
  if (format !== 'docx') {
    throw new PagekitError(`Pagekit.write does not support "${format}". Phase 1 writes "docx" only. No document was produced.`);
  }
  if (!tree || !Array.isArray(tree.blocks)) {
    throw new PagekitError('Pagekit.write expected a document tree from pagekit.parse. No document was produced.');
  }

  const ctx: WriteContext = {
    layers: resolveTemplate(options?.template),
    defaults: tree.styles ?? {},
    numbering: [],
    links: new Map(),
    nextLink: 10,
  };
  const body = renderBlocks(tree.blocks, ctx, undefined);
  const pkg = await WmlPackage.open(await seedBlob());
  const main = await pkg.mainDocumentPart();
  if (!main) throw new PagekitError('Pagekit could not open a Word package. No document was produced.');

  main.putXDocument(documentXml(body));
  const stylesPart = await main.styleDefinitionsPart();
  const numberingPart = await main.numberingDefinitionsPart();
  const settingsPart = await main.documentSettingsPart();
  if (!stylesPart || !numberingPart || !settingsPart) {
    throw new PagekitError('Pagekit could not find the styles, numbering, or settings part. No document was produced.');
  }
  stylesPart.putXDocument(stylesXml(ctx.layers));
  numberingPart.putXDocument(numberingXml(ctx.numbering));
  settingsPart.putXDocument(settingsXml());

  for (const [href, id] of ctx.links) {
    await main.addRelationship(id, HYPERLINK_REL, linkTarget(href), 'External');
  }

  const saved = await pkg.saveToBlobAsync();
  const blob = new Blob([await saved.arrayBuffer()], { type: DOCX_MIME });
  return { blob, skipped: [...(tree.skipped ?? [])] };
}

function resolveTemplate(template: DocxTemplate | undefined): TemplateLayers {
  const normal = mergeStyles(DEFAULT_NORMAL, normalizeStyleInput(template?.styles));
  const headings = {} as Record<HeadingLevel, ResolvedStyle>;
  const levels: HeadingLevel[] = [1, 2, 3, 4, 5, 6];
  for (const level of levels) {
    headings[level] = mergeStyles(DEFAULT_HEADINGS[level], normalizeStyleInput(template?.headings?.[level]));
  }
  return { normal, headings };
}

async function seedBlob(): Promise<Blob> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES);
  zip.file('_rels/.rels', PACKAGE_RELS);
  zip.file('word/document.xml', SEED_DOCUMENT);
  zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS);
  zip.file('word/styles.xml', SEED_STYLES);
  zip.file('word/numbering.xml', SEED_NUMBERING);
  zip.file('word/settings.xml', SEED_SETTINGS);
  return zip.generateAsync({ type: 'blob', mimeType: DOCX_MIME });
}

function linkTarget(href: string): string {
  try {
    return new URL(href).href;
  } catch {
    return href;
  }
}

function renderBlocks(blocks: Block[], ctx: WriteContext, extra: ResolvedStyle | undefined): XElement[] {
  const rendered: XElement[] = [];
  for (const block of blocks) rendered.push(...renderBlock(block, ctx, extra));
  return rendered;
}

function renderBlock(block: Block, ctx: WriteContext, extra: ResolvedStyle | undefined): XElement[] {
  switch (block.type) {
    case 'paragraph':
      return [renderParagraph(block, ctx, extra, null)];
    case 'list':
      return renderList(block, ctx, 0, extra);
    case 'table':
      return [renderTable(block, ctx, extra)];
    default: {
      const unknown: never = block;
      throw new PagekitError(`Unknown block type: ${JSON.stringify(unknown)}`);
    }
  }
}

function renderList(list: ListBlock, ctx: WriteContext, ilvl: number, extra: ResolvedStyle | undefined): XElement[] {
  const level = Math.max(0, Math.min(8, ilvl));
  const id = ctx.numbering.length + 1;
  ctx.numbering.push({ id, ordered: list.ordered, ilvl: level, start: list.start });
  const rendered: XElement[] = [];
  for (const item of list.items) rendered.push(...renderListItem(item, ctx, id, level, extra));
  return rendered;
}

function renderListItem(item: ListItem, ctx: WriteContext, numId: number, ilvl: number, extra: ResolvedStyle | undefined): XElement[] {
  const blocks = item.blocks.length > 0 ? item.blocks : [blankParagraph()];
  const rendered: XElement[] = [];
  let numbered = false;
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      rendered.push(renderParagraph(block, ctx, extra, numbered ? null : { numId, ilvl }));
      numbered = true;
      continue;
    }
    if (!numbered) {
      rendered.push(renderParagraph(blankParagraph(), ctx, extra, { numId, ilvl }));
      numbered = true;
    }
    if (block.type === 'list') {
      rendered.push(...renderList(block, ctx, ilvl + 1, extra));
      continue;
    }
    if (block.type === 'table') {
      rendered.push(renderTable(block, ctx, extra));
      continue;
    }
    const unknown: never = block;
    throw new PagekitError(`Unknown block in a list: ${JSON.stringify(unknown)}`);
  }
  return rendered;
}

function blankParagraph(): ParagraphBlock {
  return { type: 'paragraph', style: {}, inlines: [] };
}

function renderParagraph(
  block: ParagraphBlock,
  ctx: WriteContext,
  extra: ResolvedStyle | undefined,
  num: { numId: number; ilvl: number } | null,
): XElement {
  const template = block.headingLevel ? mergeStyles(ctx.layers.normal, ctx.layers.headings[block.headingLevel]) : ctx.layers.normal;
  const paragraphStyle = mergeStyles(paragraphStyleOf(template), paragraphStyleOf(ctx.defaults), block.style);
  const baseRun = mergeStyles(characterStyleOf(template), characterStyleOf(ctx.defaults), extra);
  const styleId = block.headingLevel ? `Heading${block.headingLevel}` : 'Normal';
  return element(
    W.p,
    paragraphProperties(paragraphStyle, {
      styleId,
      num,
      rule: block.rule === true,
      outline: block.headingLevel ? block.headingLevel - 1 : undefined,
    }),
    ...renderInlines(block.inlines, ctx, baseRun, false, block.preserveSpace === true),
  );
}

function renderInlines(
  inlines: InlineNode[],
  ctx: WriteContext,
  baseRun: ResolvedStyle,
  inLink: boolean,
  preserve: boolean,
): XElement[] {
  const rendered: XElement[] = [];
  for (const inline of inlines) {
    switch (inline.type) {
      case 'text':
        rendered.push(...textRuns(inline.text, mergeStyles(baseRun, inline.style), inLink, preserve));
        break;
      case 'lineBreak':
        rendered.push(element(W.r, runProperties(baseRun, inLink), element(W.br)));
        break;
      case 'hyperlink': {
        const linkRun = mergeStyles(baseRun, inline.style);
        const inner = renderInlines(inline.children, ctx, linkRun, true, preserve);
        if (inLink) {
          rendered.push(...inner);
          break;
        }
        const id = linkId(ctx, inline.href);
        rendered.push(element(W.hyperlink, new XAttribute(R.id, id), ...inner));
        break;
      }
      default: {
        const unknown: never = inline;
        throw new PagekitError(`Unknown inline type: ${JSON.stringify(unknown)}`);
      }
    }
  }
  return rendered;
}

function linkId(ctx: WriteContext, href: string): string {
  const existing = ctx.links.get(href);
  if (existing) return existing;
  const id = `rId${ctx.nextLink}`;
  ctx.nextLink += 1;
  ctx.links.set(href, id);
  return id;
}

function textRuns(text: string, style: ResolvedStyle, inLink: boolean, preserve: boolean): XElement[] {
  const clean = stripIllegalXml(text);
  if (!clean) return [];
  if (!preserve || !clean.includes('\n')) return [textRun(clean, style, inLink)];
  const parts = clean.split('\n');
  const runs: XElement[] = [];
  parts.forEach((part, index) => {
    if (part) runs.push(textRun(part, style, inLink));
    if (index < parts.length - 1) runs.push(element(W.r, runProperties(style, inLink), element(W.br)));
  });
  return runs;
}

function textRun(text: string, style: ResolvedStyle, inLink: boolean): XElement {
  const preserve = /^[ \t]/.test(text) || /[ \t]$/.test(text) || / {2,}/.test(text);
  const textElement = preserve
    ? element(W.t, new XAttribute(XNamespace.xml.getName('space'), 'preserve'), text)
    : element(W.t, text);
  return element(W.r, runProperties(style, inLink), textElement);
}

function runProperties(style: ResolvedStyle, inLink: boolean): XElement | null {
  const children: XElement[] = [];
  if (inLink) children.push(element(W.rStyle, new XAttribute(W.val, 'Hyperlink')));
  if (style.fontFamily) {
    children.push(element(
      W.rFonts,
      new XAttribute(W.ascii, style.fontFamily),
      new XAttribute(W.hAnsi, style.fontFamily),
      new XAttribute(W.eastAsia, style.fontFamily),
      new XAttribute(W.cs, style.fontFamily),
    ));
  }
  if (style.bold !== undefined) {
    children.push(onOff(W.b, style.bold));
    children.push(onOff(W.bCs, style.bold));
  }
  if (style.italic !== undefined) {
    children.push(onOff(W.i, style.italic));
    children.push(onOff(W.iCs, style.italic));
  }
  if (style.color) children.push(element(W.color, new XAttribute(W.val, style.color.slice(1).toUpperCase())));
  if (style.fontSizePt !== undefined) {
    const size = halfPoints(style.fontSizePt);
    children.push(element(W.sz, new XAttribute(W.val, size)));
    children.push(element(W.szCs, new XAttribute(W.val, size)));
  }
  if (style.underline !== undefined) {
    children.push(element(W.u, new XAttribute(W.val, style.underline ? 'single' : 'none')));
  }
  if (children.length === 0) return null;
  return element(W.rPr, ...children);
}

function paragraphProperties(
  style: ResolvedStyle,
  options: { styleId: string; num: { numId: number; ilvl: number } | null; rule: boolean; outline?: number },
): XElement {
  return element(
    W.pPr,
    element(W.pStyle, new XAttribute(W.val, options.styleId)),
    options.num
      ? element(
          W.numPr,
          element(W.ilvl, new XAttribute(W.val, String(options.num.ilvl))),
          element(W.numId, new XAttribute(W.val, String(options.num.numId))),
        )
      : null,
    options.rule
      ? element(
          W.pBdr,
          element(
            W.bottom,
            new XAttribute(W.val, 'single'),
            new XAttribute(W.sz, '12'),
            new XAttribute(W.space, '1'),
            new XAttribute(W.color, 'A8A29E'),
          ),
        )
      : null,
    spacingElement(style),
    style.indentLeftPt !== undefined ? element(W.ind, new XAttribute(W.left, twips(style.indentLeftPt))) : null,
    style.alignment ? element(W.jc, new XAttribute(W.val, style.alignment === 'justify' ? 'both' : style.alignment)) : null,
    options.outline !== undefined ? element(W.outlineLvl, new XAttribute(W.val, String(options.outline))) : null,
  );
}

function spacingElement(style: ResolvedStyle): XElement | null {
  const attributes: XAttribute[] = [];
  if (style.spaceBeforePt !== undefined) attributes.push(new XAttribute(W.before, twips(style.spaceBeforePt)));
  if (style.spaceAfterPt !== undefined) attributes.push(new XAttribute(W.after, twips(style.spaceAfterPt)));
  if (style.lineSpacingPt !== undefined) {
    attributes.push(new XAttribute(W.line, twips(style.lineSpacingPt)));
    attributes.push(new XAttribute(W.lineRule, 'exact'));
  }
  if (attributes.length === 0) return null;
  return element(W.spacing, ...attributes);
}

function renderTable(table: TableBlock, ctx: WriteContext, extra: ResolvedStyle | undefined): XElement {
  const columns = Math.max(1, ...table.rows.map((row) => row.cells.length));
  const widths = columnWidths(columns);
  return element(
    W.tbl,
    element(
      W.tblPr,
      element(W.tblW, new XAttribute(W._w, String(CONTENT_WIDTH)), new XAttribute(W.type, 'dxa')),
      element(W.tblBorders, ...borderSet('top', 'left', 'bottom', 'right', 'insideH', 'insideV')),
      element(W.tblLayout, new XAttribute(W.type, 'fixed')),
    ),
    element(W.tblGrid, ...widths.map((width) => element(W.gridCol, new XAttribute(W._w, String(width))))),
    ...table.rows.map((row) => element(
      W.tr,
      ...row.cells.map((cell, index) => {
        const width = widths[index] ?? widths[widths.length - 1] ?? 1000;
        return element(
          W.tc,
          element(
            W.tcPr,
            element(W.tcW, new XAttribute(W._w, String(width)), new XAttribute(W.type, 'dxa')),
            cell.header
              ? element(W.shd, new XAttribute(W.val, 'clear'), new XAttribute(W.color, 'auto'), new XAttribute(W.fill, 'E7F0EA'))
              : null,
          ),
          ...renderBlocks(cell.blocks, ctx, cell.header ? mergeStyles(extra, { bold: true }) : extra),
        );
      }),
    )),
  );
}

function columnWidths(count: number): number[] {
  const base = Math.floor(CONTENT_WIDTH / count);
  const widths = Array.from({ length: count }, () => base);
  const last = widths.length - 1;
  if (last >= 0) widths[last] = CONTENT_WIDTH - base * last;
  return widths;
}

const BORDER_NAMES = {
  top: W.top,
  left: W.left,
  bottom: W.bottom,
  right: W.right,
  insideH: W.insideH,
  insideV: W.insideV,
} as const;

function borderSet(...names: Array<keyof typeof BORDER_NAMES>): XElement[] {
  return names.map((name) => element(
    BORDER_NAMES[name],
    new XAttribute(W.val, 'single'),
    new XAttribute(W.sz, '4'),
    new XAttribute(W.space, '0'),
    new XAttribute(W.color, 'C4C0B8'),
  ));
}

function documentXml(blocks: XElement[]): XDocument {
  return new XDocument(
    new XDeclaration('1.0', 'UTF-8', 'yes'),
    element(W.document, namespaceAttr('w'), namespaceAttr('r'), element(W.body, ...blocks, sectionProperties())),
  );
}

function sectionProperties(): XElement {
  return element(
    W.sectPr,
    element(W.pgSz, new XAttribute(W._w, '12240'), new XAttribute(W.h, '15840')),
    element(
      W.pgMar,
      new XAttribute(W.top, '1440'),
      new XAttribute(W.right, '1440'),
      new XAttribute(W.bottom, '1440'),
      new XAttribute(W.left, '1440'),
      new XAttribute(W.header, '720'),
      new XAttribute(W.footer, '720'),
      new XAttribute(W.gutter, '0'),
    ),
  );
}

function stylesXml(layers: TemplateLayers): XDocument {
  const levels: HeadingLevel[] = [1, 2, 3, 4, 5, 6];
  return new XDocument(
    new XDeclaration('1.0', 'UTF-8', 'yes'),
    element(
      W.styles,
      namespaceAttr('w'),
      element(
        W.docDefaults,
        element(W.rPrDefault, runProperties(layers.normal, false)),
        element(W.pPrDefault, element(W.pPr, spacingElement(layers.normal))),
      ),
      styleDefinition('Normal', 'Normal', 'paragraph', true, null, layers.normal, null),
      ...levels.map((level) => styleDefinition(
        `Heading${level}`,
        `heading ${level}`,
        'paragraph',
        false,
        'Normal',
        mergeStyles(layers.normal, layers.headings[level]),
        level - 1,
      )),
      element(
        W.style,
        new XAttribute(W.type, 'character'),
        new XAttribute(W.styleId, 'Hyperlink'),
        element(W_NAME, new XAttribute(W.val, 'Hyperlink')),
        element(W.uiPriority, new XAttribute(W.val, '99')),
        runProperties({ color: '#0563C1', underline: true }, false),
      ),
    ),
  );
}

function styleDefinition(
  id: string,
  name: string,
  type: 'paragraph' | 'character',
  isDefault: boolean,
  basedOn: string | null,
  style: ResolvedStyle,
  outline: number | null,
): XElement {
  const paragraph = element(
    W.pPr,
    spacingElement(paragraphStyleOf(style)),
    style.alignment ? element(W.jc, new XAttribute(W.val, style.alignment === 'justify' ? 'both' : style.alignment)) : null,
    outline !== null ? element(W.outlineLvl, new XAttribute(W.val, String(outline))) : null,
  );
  return element(
    W.style,
    new XAttribute(W.type, type),
    new XAttribute(W.styleId, id),
    isDefault ? new XAttribute(W._default, '1') : null,
    element(W_NAME, new XAttribute(W.val, name)),
    basedOn ? element(W.basedOn, new XAttribute(W.val, basedOn)) : null,
    type === 'paragraph' ? element(W.next, new XAttribute(W.val, 'Normal')) : null,
    element(W.uiPriority, new XAttribute(W.val, id.startsWith('Heading') ? String(8 + (outline ?? 0)) : '0')),
    element(W.qFormat),
    type === 'paragraph' ? paragraph : null,
    runProperties(characterStyleOf(style), false),
  );
}

function numberingXml(lists: ListUse[]): XDocument {
  return new XDocument(
    new XDeclaration('1.0', 'UTF-8', 'yes'),
    element(
      W.numbering,
      namespaceAttr('w'),
      abstractNumbering(0, false),
      abstractNumbering(1, true),
      ...lists.map((list) => element(
        W.num,
        new XAttribute(W.numId, String(list.id)),
        element(W.abstractNumId, new XAttribute(W.val, list.ordered ? '1' : '0')),
        list.start === 1
          ? null
          : element(
              W.lvlOverride,
              new XAttribute(W.ilvl, String(list.ilvl)),
              element(W.startOverride, new XAttribute(W.val, String(list.start))),
            ),
      )),
    ),
  );
}

function abstractNumbering(id: number, ordered: boolean): XElement {
  const levels = Array.from({ length: 9 }, (_, level) => levelElement(level, ordered));
  return element(
    W.abstractNum,
    new XAttribute(W.abstractNumId, String(id)),
    element(W.nsid, new XAttribute(W.val, id === 0 ? 'A1B2C3D4' : 'B2C3D4E5')),
    element(W.multiLevelType, new XAttribute(W.val, 'hybridMultilevel')),
    ...levels,
  );
}

function levelElement(level: number, ordered: boolean): XElement {
  const bullet = BULLETS[level % BULLETS.length] ?? BULLETS[0];
  const left = 720 * (level + 1);
  return element(
    W.lvl,
    new XAttribute(W.ilvl, String(level)),
    element(W.start, new XAttribute(W.val, '1')),
    element(W.numFmt, new XAttribute(W.val, ordered ? 'decimal' : 'bullet')),
    element(W.lvlText, new XAttribute(W.val, ordered ? `%${level + 1}.` : bullet.text)),
    element(W.lvlJc, new XAttribute(W.val, 'left')),
    element(W.pPr, element(W.ind, new XAttribute(W.left, String(left)), new XAttribute(W.hanging, '360'))),
    ordered
      ? null
      : element(
          W.rPr,
          element(
            W.rFonts,
            new XAttribute(W.ascii, bullet.font),
            new XAttribute(W.hAnsi, bullet.font),
            new XAttribute(W.hint, 'default'),
          ),
        ),
  );
}

function settingsXml(): XDocument {
  return new XDocument(new XDeclaration('1.0', 'UTF-8', 'yes'), element(W.settings, namespaceAttr('w')));
}

function namespaceAttr(prefix: 'w' | 'r'): XAttribute {
  const uri = prefix === 'w'
    ? 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  return new XAttribute(XNamespace.xmlns.getName(prefix), uri);
}

function onOff(name: XName, value: boolean): XElement {
  return value ? element(name) : element(name, new XAttribute(W.val, '0'));
}

function element(name: XName, ...content: Array<XElement | XAttribute | string | null | undefined | false>): XElement {
  const nodes = content.filter((node): node is XElement | XAttribute | string => node !== null && node !== undefined && node !== false);
  return new XElement(name, ...nodes);
}
