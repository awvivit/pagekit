import { parseFragment, type DefaultTreeAdapterMap } from 'parse5';
import { PagekitError } from './errors.js';
import type {
  Block,
  DocumentTree,
  HeadingLevel,
  InlineNode,
  ListBlock,
  ParagraphBlock,
  ParseOptions,
  ResolvedStyle,
  SkippedItem,
  TableBlock,
  TableCell,
} from './types.js';
import {
  cssColorToHex,
  cssLengthToPt,
  firstFontFamily,
  normalizeAlignment,
  normalizeStyleInput,
  stripIllegalXml,
} from './units.js';

type P5Child = DefaultTreeAdapterMap['childNode'];

type ParseContext = {
  skipped: SkippedItem[];
  specified: WeakMap<Element, Set<string>>;
};

type CursorItem =
  | { kind: 'text'; text: string }
  | { kind: 'comment' }
  | { kind: 'element'; live: Element };

const BLOCK_TAGS = new Set([
  'address',
  'article',
  'aside',
  'blockquote',
  'center',
  'dd',
  'details',
  'dialog',
  'div',
  'dl',
  'dt',
  'fieldset',
  'figcaption',
  'figure',
  'footer',
  'form',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'header',
  'hgroup',
  'hr',
  'li',
  'main',
  'nav',
  'ol',
  'p',
  'pre',
  'section',
  'table',
  'ul',
]);

const INHERITED_PROPS = new Set([
  'color',
  'font-family',
  'font-size',
  'font-weight',
  'font-style',
  'text-align',
  'line-height',
  'letter-spacing',
  'white-space',
]);

const AUTHOR_PROPS = [
  'font-family',
  'font-size',
  'font-weight',
  'font-style',
  'color',
  'text-align',
  'text-decoration-line',
  'margin-top',
  'margin-bottom',
  'margin-left',
  'line-height',
  'break-before',
  'break-after',
  'page-break-before',
  'page-break-after',
];

/**
 * Read an element (or an element id) into a format-neutral tree.
 * Structure comes from parse5 so the input is never re-inserted as HTML.
 * Styles come from computed CSS. Element CSS is stored on each node and wins
 * over `options.styles`. The template is applied later, by `write`.
 */
export function parse(rootOrId: Element | string, options?: ParseOptions): DocumentTree {
  const root = resolveRoot(rootOrId);
  const ctx: ParseContext = { skipped: [], specified: new WeakMap() };
  return {
    styles: normalizeStyleInput(options?.styles),
    blocks: convertElement(root, ctx),
    skipped: ctx.skipped,
  };
}

function resolveRoot(rootOrId: Element | string): Element {
  if (typeof rootOrId === 'string') {
    if (typeof document === 'undefined') {
      throw new PagekitError('Pagekit.parse needs a browser document to look up an element id. No document was produced.');
    }
    const found = document.getElementById(rootOrId);
    if (!found) {
      throw new PagekitError(`No element with id "${rootOrId}" was found. No document was produced.`);
    }
    return found;
  }
  if (!rootOrId || rootOrId.nodeType !== 1) {
    throw new PagekitError('Pagekit.parse expected an element or an element id string. No document was produced.');
  }
  return rootOrId;
}

function skip(ctx: ParseContext, kind: string, message: string, tag?: string): void {
  const item: SkippedItem = { kind, message };
  if (tag) item.tag = tag;
  ctx.skipped.push(item);
}

function computedStyle(el: Element): CSSStyleDeclaration {
  const view = el.ownerDocument.defaultView;
  if (!view) {
    throw new PagekitError('Pagekit.parse needs a document window to read computed styles. No document was produced.');
  }
  return view.getComputedStyle(el);
}

function inlineStyle(el: Element): CSSStyleDeclaration | null {
  if (!('style' in el)) return null;
  const style: unknown = el.style;
  if (typeof style !== 'object' || style === null || !('getPropertyValue' in style)) return null;
  return style as CSSStyleDeclaration;
}

function collectDeclared(style: CSSStyleDeclaration, props: Set<string>): void {
  for (let index = 0; index < style.length; index += 1) {
    const name = style.item(index);
    if (name) props.add(name);
  }
  for (const name of AUTHOR_PROPS) {
    if (style.getPropertyValue(name)) props.add(name);
  }
}

function visitRules(rules: CSSRuleList, el: Element, props: Set<string>): void {
  for (const rule of Array.from(rules)) {
    if (rule instanceof CSSStyleRule) {
      if (selectorMatches(el, rule.selectorText)) collectDeclared(rule.style, props);
      continue;
    }
    if ('cssRules' in rule && rule.cssRules instanceof CSSRuleList) {
      visitRules(rule.cssRules, el, props);
    }
  }
}

function specifiedProps(el: Element, ctx: ParseContext): Set<string> {
  const cached = ctx.specified.get(el);
  if (cached) return cached;
  const props = new Set<string>();
  const inline = inlineStyle(el);
  if (inline) collectDeclared(inline, props);
  const sheets = el.ownerDocument.styleSheets;
  for (const sheet of Array.from(sheets)) {
    if (!(sheet.ownerNode instanceof Element)) continue;
    let rules: CSSRuleList;
    try {
      rules = sheet.cssRules;
    } catch {
      continue;
    }
    visitRules(rules, el, props);
  }
  ctx.specified.set(el, props);
  return props;
}

function authorWins(el: Element, prop: string, ctx: ParseContext): boolean {
  if (specifiedProps(el, ctx).has(prop)) return true;
  if (!INHERITED_PROPS.has(prop)) return false;
  const parent = el.parentElement;
  if (!parent || !authorWins(parent, prop, ctx)) return false;
  // A user-agent rule on this element (for example a link's blue) beats inheritance.
  // When that happens the computed value no longer matches the parent.
  return computedStyle(el).getPropertyValue(prop) === computedStyle(parent).getPropertyValue(prop);
}

function declaredValue(el: Element, prop: string): string {
  const inline = inlineStyle(el)?.getPropertyValue(prop);
  if (inline) return inline;
  return computedStyle(el).getPropertyValue(prop);
}

function isPageBreak(value: string): boolean {
  return value === 'page' || value === 'always' || value === 'left' || value === 'right';
}

function notePageBreak(el: Element, ctx: ParseContext): void {
  const props = ['break-before', 'break-after', 'page-break-before', 'page-break-after'];
  const broken = props.some((prop) => specifiedProps(el, ctx).has(prop) && isPageBreak(declaredValue(el, prop)));
  if (!broken) return;
  skip(
    ctx,
    'page-break',
    'Page breaks are not supported. The surrounding text was kept without a break.',
    el.tagName.toLowerCase(),
  );
}

function noteLayout(el: Element, ctx: ParseContext): void {
  const display = computedStyle(el).display;
  if (display === 'flex' || display === 'inline-flex') {
    skip(ctx, 'flex', 'Flex layout is not supported. The text was kept in document order.', el.tagName.toLowerCase());
  } else if (display === 'grid' || display === 'inline-grid') {
    skip(ctx, 'grid', 'Grid layout is not supported. The text was kept in document order.', el.tagName.toLowerCase());
  }
}

function characterFromCss(el: Element, ctx: ParseContext): ResolvedStyle {
  const style: ResolvedStyle = {};
  const computed = computedStyle(el);
  if (authorWins(el, 'font-family', ctx)) {
    const family = firstFontFamily(computed.fontFamily);
    if (family) style.fontFamily = family;
  }
  if (authorWins(el, 'font-size', ctx)) {
    const size = cssLengthToPt(computed.fontSize);
    if (size !== undefined && size > 0) style.fontSizePt = size;
  }
  if (authorWins(el, 'color', ctx)) {
    const color = cssColorToHex(computed.color);
    if (color) style.color = color;
  }
  if (authorWins(el, 'font-weight', ctx)) {
    const weight = computed.fontWeight.trim().toLowerCase();
    const numeric = Number(weight);
    style.bold = weight === 'bold' || weight === 'bolder' || (Number.isFinite(numeric) && numeric >= 600);
  }
  if (authorWins(el, 'font-style', ctx)) {
    const fontStyle = computed.fontStyle.trim().toLowerCase();
    style.italic = fontStyle === 'italic' || fontStyle === 'oblique';
  }
  if (authorWins(el, 'text-decoration-line', ctx) || specifiedProps(el, ctx).has('text-decoration')) {
    const line = (computed.textDecorationLine || computed.textDecoration || '').toLowerCase();
    style.underline = line.split(/\s+/).includes('underline');
  }
  return style;
}

function paragraphFromCss(el: Element, ctx: ParseContext): ResolvedStyle {
  const style: ResolvedStyle = {};
  const computed = computedStyle(el);
  if (authorWins(el, 'text-align', ctx)) {
    const alignment = normalizeAlignment(computed.textAlign);
    if (alignment) style.alignment = alignment;
  }
  const fontSize = cssLengthToPt(computed.fontSize) ?? 12;
  if (authorWins(el, 'margin-top', ctx)) {
    const before = cssLengthToPt(computed.marginTop, fontSize);
    if (before !== undefined) style.spaceBeforePt = before;
  }
  if (authorWins(el, 'margin-bottom', ctx)) {
    const after = cssLengthToPt(computed.marginBottom, fontSize);
    if (after !== undefined) style.spaceAfterPt = after;
  }
  if (authorWins(el, 'margin-left', ctx)) {
    const indent = cssLengthToPt(computed.marginLeft, fontSize);
    if (indent !== undefined) style.indentLeftPt = indent;
  }
  if (authorWins(el, 'line-height', ctx)) {
    const raw = computed.lineHeight.trim().toLowerCase();
    if (raw && raw !== 'normal') {
      const line = /^[0-9.]+$/.test(raw) ? fontSize * Number(raw) : cssLengthToPt(raw, fontSize);
      if (line !== undefined && line > 0) style.lineSpacingPt = line;
    }
  }
  return style;
}

function semanticCharacter(el: Element): ResolvedStyle {
  const tag = el.tagName.toLowerCase();
  const style: ResolvedStyle = {};
  if (tag === 'b' || tag === 'strong') style.bold = true;
  if (tag === 'i' || tag === 'em') style.italic = true;
  if (tag === 'u') style.underline = true;
  if (tag === 'a') {
    style.underline = true;
    style.color = '#0563C1';
  }
  return style;
}

function runStyle(el: Element, ctx: ParseContext): ResolvedStyle {
  const style: ResolvedStyle = { ...semanticCharacter(el) };
  const authored = characterFromCss(el, ctx);
  for (const [key, value] of Object.entries(authored)) {
    if (value !== undefined) Object.assign(style, { [key]: value });
  }
  return style;
}

function preservesWhitespace(el: Element): boolean {
  const tag = el.tagName.toLowerCase();
  if (tag === 'pre') return true;
  const value = computedStyle(el).whiteSpace;
  return value === 'pre' || value === 'pre-wrap' || value === 'break-spaces';
}

function headingLevelOf(tag: string): HeadingLevel | undefined {
  switch (tag) {
    case 'h1':
      return 1;
    case 'h2':
      return 2;
    case 'h3':
      return 3;
    case 'h4':
      return 4;
    case 'h5':
      return 5;
    case 'h6':
      return 6;
    default:
      return undefined;
  }
}

function isBlockish(el: Element): boolean {
  const tag = el.tagName.toLowerCase();
  if (BLOCK_TAGS.has(tag)) return true;
  const display = computedStyle(el).display;
  return (
    display === 'block' ||
    display === 'flex' ||
    display === 'grid' ||
    display === 'list-item' ||
    display === 'table' ||
    display === 'flow-root'
  );
}

function childItems(live: Element): CursorItem[] {
  const fromParser = itemsFromParse5(live);
  return fromParser ?? itemsFromLive(live);
}

function itemsFromParse5(live: Element): CursorItem[] | null {
  const fragment = parseFragment(live.innerHTML);
  const elements = [...live.children];
  let index = 0;
  const items: CursorItem[] = [];
  for (const node of fragment.childNodes) {
    const item = cursorItem(node, elements, index);
    if (item === 'mismatch') return null;
    if (item.kind === 'element') index += 1;
    items.push(item.item);
  }
  if (index !== elements.length) return null;
  return items;
}

function cursorItem(
  node: P5Child,
  elements: Element[],
  index: number,
): { kind: 'element'; item: CursorItem } | { kind: 'other'; item: CursorItem } | 'mismatch' {
  if (node.nodeName === '#text' && 'value' in node) {
    return { kind: 'other', item: { kind: 'text', text: node.value } };
  }
  if (node.nodeName === '#comment') {
    return { kind: 'other', item: { kind: 'comment' } };
  }
  if (node.nodeName === '#documentType') {
    return { kind: 'other', item: { kind: 'text', text: '' } };
  }
  const child = elements[index];
  if (!child || child.tagName.toLowerCase() !== node.nodeName) return 'mismatch';
  return { kind: 'element', item: { kind: 'element', live: child } };
}

function itemsFromLive(live: Element): CursorItem[] {
  const items: CursorItem[] = [];
  for (const node of Array.from(live.childNodes)) {
    if (node.nodeType === Node.TEXT_NODE) {
      items.push({ kind: 'text', text: node.nodeValue ?? '' });
    } else if (node.nodeType === Node.COMMENT_NODE) {
      items.push({ kind: 'comment' });
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      items.push({ kind: 'element', live: node as Element });
    }
  }
  return items;
}

function selectorMatches(el: Element, selector: string): boolean {
  try {
    return el.matches(selector);
  } catch {
    return false;
  }
}

function cleanText(text: string): string {
  return stripIllegalXml(text);
}

function normalizeText(text: string, preserve: boolean): string {
  const clean = cleanText(text);
  if (preserve) return clean.replace(/\r\n?/g, '\n');
  return clean.replace(/[ \t\n\f\v\r]+/g, ' ');
}

function paragraph(el: Element, ctx: ParseContext, inlines: InlineNode[], preserve: boolean): ParagraphBlock {
  return {
    type: 'paragraph',
    style: paragraphFromCss(el, ctx),
    inlines,
    preserveSpace: preserve,
  };
}

function emptyParagraph(el: Element, ctx: ParseContext, preserve = false): ParagraphBlock {
  return paragraph(el, ctx, [], preserve);
}

function visibleText(el: Element): string {
  return cleanText(el.textContent ?? '').replace(/\s+/g, ' ').trim();
}

function formControlText(el: Element): string {
  const tag = el.tagName.toLowerCase();
  if (tag === 'input') {
    const type = (el.getAttribute('type') || 'text').toLowerCase();
    if (type === 'hidden' || type === 'password' || type === 'file' || type === 'checkbox' || type === 'radio' || type === 'button' || type === 'submit' || type === 'reset') {
      return type === 'button' || type === 'submit' || type === 'reset' ? (el.getAttribute('value') || '').trim() : '';
    }
    return (el.getAttribute('value') || '').trim();
  }
  if (tag === 'textarea' || tag === 'button') return visibleText(el);
  if (tag === 'select') {
    return [...el.querySelectorAll('option')].map((option) => (option.textContent ?? '').trim()).filter(Boolean).join(', ');
  }
  return '';
}

function unsafeHref(href: string): boolean {
  const lower = href.trim().toLowerCase();
  return lower.startsWith('javascript:') || lower.startsWith('vbscript:') || lower.startsWith('data:');
}

function convertElement(el: Element, ctx: ParseContext): Block[] {
  const tag = el.tagName.toLowerCase();
  notePageBreak(el, ctx);
  noteLayout(el, ctx);

  if (tag === 'script') {
    skip(ctx, 'script', 'Scripts are not run and are not included in the document.', tag);
    return [];
  }
  if (tag === 'style') {
    skip(ctx, 'style', 'Stylesheets are not copied into the document as text.', tag);
    return [];
  }
  if (tag === 'template') {
    skip(ctx, 'template', 'Template contents are not rendered. Their text was kept as text.', tag);
    return textParagraph(el, ctx, visibleText(el));
  }
  if (tag === 'noscript') {
    skip(ctx, 'noscript', 'noscript content was kept as text and was not executed.', tag);
    return textParagraph(el, ctx, visibleText(el));
  }
  if (tag === 'iframe') {
    skip(ctx, 'iframe', 'Embedded frames are not supported. Their text was kept as text.', tag);
    return textParagraph(el, ctx, el.getAttribute('title') || visibleText(el));
  }
  if (tag === 'object') {
    skip(ctx, 'object', 'Embedded objects are not supported. Their text was kept as text.', tag);
    return textParagraph(el, ctx, visibleText(el));
  }
  if (tag === 'embed') {
    skip(ctx, 'embed', 'Embedded objects are not supported. Their text was kept as text.', tag);
    return [];
  }
  if (tag === 'svg') {
    skip(ctx, 'svg', 'SVG graphics are not embedded. Their text was kept as text.', tag);
    return textParagraph(el, ctx, visibleText(el));
  }
  if (tag === 'canvas') {
    skip(ctx, 'canvas', 'Canvas graphics are not embedded. Any label was kept as text.', tag);
    return textParagraph(el, ctx, el.getAttribute('aria-label') || el.getAttribute('title') || '');
  }
  if (tag === 'video' || tag === 'audio') {
    skip(ctx, 'media', 'Audio and video are not embedded. Their text was kept as text.', tag);
    return textParagraph(el, ctx, visibleText(el));
  }
  if (tag === 'img') {
    return imageBlocks(el, ctx);
  }
  if (isFormControl(tag)) {
    skip(ctx, 'form', `Form control <${tag}> is not supported. Its text was kept as text.`, tag);
    return textParagraph(el, ctx, formControlText(el));
  }
  if (tag === 'ul' || tag === 'ol') return [convertList(el, ctx)];
  if (tag === 'table') return convertTable(el, ctx);
  if (tag === 'hr') {
    return [{ ...emptyParagraph(el, ctx), rule: true }];
  }
  if (tag === 'header' || tag === 'footer') {
    skip(
      ctx,
      tag,
      `HTML <${tag}> elements are kept in the document body. Word ${tag}s are not created.`,
      tag,
    );
  }
  if (tag === 'form' || tag === 'fieldset') {
    skip(ctx, 'form', `Form element <${tag}> is not supported. Its text was kept as text.`, tag);
  }

  const level = headingLevelOf(tag);
  if (level) return convertHeading(el, ctx, level);
  if (isBlockish(el) || tag === 'li') return convertBlockContainer(el, ctx, { force: tag === 'p' || tag === 'pre' || tag === 'li' });

  const inlines = convertInline(el, ctx);
  if (inlines.length === 0) return [];
  return [paragraph(el, ctx, trimInlines(inlines, preservesWhitespace(el)), preservesWhitespace(el))];
}

function isFormControl(tag: string): boolean {
  return tag === 'input' || tag === 'textarea' || tag === 'select' || tag === 'button' || tag === 'progress' || tag === 'meter' || tag === 'output' || tag === 'datalist';
}

function textParagraph(el: Element, ctx: ParseContext, text: string): Block[] {
  const trimmed = text.trim();
  if (!trimmed) return [];
  return [paragraph(el, ctx, [{ type: 'text', text: trimmed, style: runStyle(el, ctx) }], false)];
}

function imageBlocks(el: Element, ctx: ParseContext): Block[] {
  skip(ctx, 'image', 'Images are not embedded. The alt text was kept as text.', 'img');
  return textParagraph(el, ctx, el.getAttribute('alt') || '');
}

function convertHeading(el: Element, ctx: ParseContext, level: HeadingLevel): Block[] {
  const blocks = convertBlockContainer(el, ctx, { force: true });
  const [first, ...rest] = blocks;
  if (!first) return [{ ...emptyParagraph(el, ctx), headingLevel: level }];
  if (first.type === 'paragraph') {
    return [{ ...first, headingLevel: level }, ...rest];
  }
  return [{ ...emptyParagraph(el, ctx), headingLevel: level }, ...blocks];
}

function convertBlockContainer(el: Element, ctx: ParseContext, options?: { force?: boolean }): Block[] {
  const preserve = preservesWhitespace(el);
  const blocks: Block[] = [];
  let pending: InlineNode[] = [];
  const flush = (): void => {
    const inlines = trimInlines(pending, preserve);
    pending = [];
    if (inlines.length === 0) return;
    blocks.push(paragraph(el, ctx, inlines, preserve));
  };
  for (const item of childItems(el)) {
    if (item.kind === 'comment') {
      skip(ctx, 'comment', 'HTML comments are not written into the document.');
      continue;
    }
    if (item.kind === 'text') {
      const text = normalizeText(item.text, preserve);
      if (!text) continue;
      if (!preserve && text === ' ' && pending.length === 0) continue;
      pending.push({ type: 'text', text, style: runStyle(el, ctx) });
      continue;
    }
    const child = item.live;
    if (child.tagName.toLowerCase() === 'br') {
      pending.push({ type: 'lineBreak' });
      continue;
    }
    if (isBlockish(child) || child.tagName.toLowerCase() === 'img') {
      flush();
      blocks.push(...convertElement(child, ctx));
      continue;
    }
    pending.push(...convertInline(child, ctx));
  }
  flush();
  if (options?.force && !blocks.some((block) => block.type === 'paragraph')) {
    return [emptyParagraph(el, ctx, preserve), ...blocks];
  }
  return blocks;
}

function convertInline(el: Element, ctx: ParseContext): InlineNode[] {
  const tag = el.tagName.toLowerCase();
  notePageBreak(el, ctx);
  if (tag === 'br') return [{ type: 'lineBreak' }];
  if (tag === 'script' || tag === 'style' || tag === 'template' || tag === 'img' || isFormControl(tag) || tag === 'svg' || tag === 'canvas' || tag === 'video' || tag === 'audio' || tag === 'iframe' || tag === 'object' || tag === 'embed') {
    const blocks = convertElement(el, ctx);
    return blocksToInlines(blocks);
  }
  if (tag === 'del' || tag === 'ins') {
    skip(ctx, 'tracked-change', 'Tracked changes are not supported. The text was kept as ordinary text.', tag);
  }
  const style = runStyle(el, ctx);
  if (tag === 'a') {
    const href = (el.getAttribute('href') || '').trim();
    const children = inlineChildren(el, ctx, style);
    if (!href) return children;
    if (unsafeHref(href)) {
      skip(ctx, 'unsafe-link', 'The link was not written because its address is not a normal web link. The link text was kept.', 'a');
      return children;
    }
    const content = children.length > 0 ? children : [{ type: 'text' as const, text: href, style }];
    return [{ type: 'hyperlink', href, style, children: content }];
  }
  return inlineChildren(el, ctx, style);
}

function inlineChildren(el: Element, ctx: ParseContext, style: ResolvedStyle): InlineNode[] {
  const preserve = preservesWhitespace(el);
  const inlines: InlineNode[] = [];
  for (const item of childItems(el)) {
    if (item.kind === 'comment') {
      skip(ctx, 'comment', 'HTML comments are not written into the document.');
      continue;
    }
    if (item.kind === 'text') {
      const text = normalizeText(item.text, preserve);
      if (!text) continue;
      inlines.push({ type: 'text', text, style });
      continue;
    }
    const child = item.live;
    if (isBlockish(child)) {
      inlines.push(...blocksToInlines(convertElement(child, ctx)));
      continue;
    }
    inlines.push(...convertInline(child, ctx));
  }
  return inlines;
}

function blocksToInlines(blocks: Block[]): InlineNode[] {
  const inlines: InlineNode[] = [];
  for (const block of blocks) {
    if (inlines.length > 0) inlines.push({ type: 'lineBreak' });
    if (block.type === 'paragraph') {
      inlines.push(...block.inlines);
    } else if (block.type === 'list' || block.type === 'table') {
      inlines.push(...blocksToInlines(block.type === 'list' ? block.items.flatMap((item) => item.blocks) : block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks))));
    } else {
      const unknown: never = block;
      throw new PagekitError(`Unknown block while reading inline content: ${JSON.stringify(unknown)}`);
    }
  }
  return inlines;
}

function trimInlines(inlines: InlineNode[], preserve: boolean): InlineNode[] {
  const copy = inlines.slice();
  if (preserve) return copy.filter((inline) => inline.type !== 'text' || inline.text.length > 0);
  while (copy.length > 0) {
    const first = copy[0];
    if (first?.type !== 'text') break;
    first.text = first.text.replace(/^[ \t\n\f\v\r]+/, '');
    if (first.text.length > 0) break;
    copy.shift();
  }
  while (copy.length > 0) {
    const last = copy[copy.length - 1];
    if (last?.type !== 'text') break;
    last.text = last.text.replace(/[ \t\n\f\v\r]+$/, '');
    if (last.text.length > 0) break;
    copy.pop();
  }
  return copy.filter((inline) => inline.type !== 'text' || inline.text.length > 0);
}

function convertList(el: Element, ctx: ParseContext): ListBlock {
  const ordered = el.tagName.toLowerCase() === 'ol';
  const startRaw = Number(el.getAttribute('start') ?? '1');
  const start = Number.isFinite(startRaw) ? startRaw : 1;
  const items = [...el.children].filter((child) => child.tagName.toLowerCase() === 'li');
  return {
    type: 'list',
    ordered,
    start,
    items: items.map((item) => {
      const blocks = convertBlockContainer(item, ctx, { force: true });
      return { blocks: blocks.length > 0 ? blocks : [emptyParagraph(item, ctx)] };
    }),
  };
}

function convertTable(el: Element, ctx: ParseContext): Block[] {
  const before: Block[] = [];
  const rows: Element[] = [];
  for (const child of Array.from(el.children)) {
    const tag = child.tagName.toLowerCase();
    if (tag === 'caption') {
      before.push(...convertBlockContainer(child, ctx, { force: true }));
    } else if (tag === 'tr') {
      rows.push(child);
    } else if (tag === 'thead' || tag === 'tbody' || tag === 'tfoot') {
      for (const row of Array.from(child.children)) {
        if (row.tagName.toLowerCase() === 'tr') rows.push(row);
      }
    }
  }
  const parsedRows = rows.map((row) => cellsOf(row));
  const merged = parsedRows.some((row) => row.some((cell) => cell.span > 1));
  if (merged) {
    skip(ctx, 'merged-cells', 'Merged table cells are not supported. The cell text was kept as paragraphs.', 'table');
    const flattened: Block[] = [...before];
    for (const row of parsedRows) {
      for (const cell of row) flattened.push(...convertBlockContainer(cell.el, ctx, { force: true }));
    }
    return flattened;
  }
  const table: TableBlock = {
    type: 'table',
    rows: parsedRows.map((row) => ({
      cells: row.map((cell): TableCell => {
        const blocks = convertBlockContainer(cell.el, ctx, { force: true });
        return {
          header: cell.el.tagName.toLowerCase() === 'th',
          blocks: blocks.length > 0 ? blocks : [emptyParagraph(cell.el, ctx)],
        };
      }),
    })),
  };
  if (table.rows.length === 0) return before;
  return [...before, table];
}

function cellsOf(row: Element): Array<{ el: Element; span: number }> {
  const cells: Array<{ el: Element; span: number }> = [];
  for (const child of Array.from(row.children)) {
    const tag = child.tagName.toLowerCase();
    if (tag !== 'td' && tag !== 'th') continue;
    const colspan = Number(child.getAttribute('colspan') || '1');
    const rowspan = Number(child.getAttribute('rowspan') || '1');
    const span = Math.max(Number.isFinite(colspan) ? colspan : 1, Number.isFinite(rowspan) ? rowspan : 1);
    cells.push({ el: child, span });
  }
  return cells;
}
