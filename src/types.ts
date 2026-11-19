/** Horizontal alignment stored on the document tree and written to Word. */
export type Alignment = 'left' | 'center' | 'right' | 'justify';

export type HeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;

/**
 * Styles after units are resolved.
 * Lengths are points. Colors are `#RRGGBB`. Absent fields are not set at this layer.
 */
export type ResolvedStyle = {
  fontFamily?: string;
  fontSizePt?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  alignment?: Alignment;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  lineSpacingPt?: number;
  indentLeftPt?: number;
};

/**
 * Style values accepted from callers.
 * Numbers are points. Strings use CSS units (`12pt`, `16px`) or CSS colors.
 */
export type StyleInput = {
  fontFamily?: string;
  fontSize?: number | string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  alignment?: Alignment | 'start' | 'end';
  spaceBefore?: number | string;
  spaceAfter?: number | string;
  lineSpacing?: number | string;
  indentLeft?: number | string;
};

/** Something the writer could not represent. Visible text is still kept. */
export type SkippedItem = {
  kind: string;
  message: string;
  tag?: string;
};

export type TextInline = {
  type: 'text';
  text: string;
  style: ResolvedStyle;
};

export type LineBreakInline = {
  type: 'lineBreak';
};

export type HyperlinkInline = {
  type: 'hyperlink';
  href: string;
  style: ResolvedStyle;
  children: InlineNode[];
};

export type InlineNode = TextInline | LineBreakInline | HyperlinkInline;

export type ParagraphBlock = {
  type: 'paragraph';
  style: ResolvedStyle;
  inlines: InlineNode[];
  headingLevel?: HeadingLevel;
  preserveSpace?: boolean;
  /** Drawn as a bottom border. Used for horizontal rules. */
  rule?: boolean;
};

export type ListItem = {
  blocks: Block[];
};

export type ListBlock = {
  type: 'list';
  ordered: boolean;
  start: number;
  items: ListItem[];
};

export type TableCell = {
  header: boolean;
  blocks: Block[];
};

export type TableRow = {
  cells: TableCell[];
};

export type TableBlock = {
  type: 'table';
  rows: TableRow[];
};

export type Block = ParagraphBlock | ListBlock | TableBlock;

/**
 * Format-neutral document.
 * `styles` is the caller's style object (middle cascade layer), already in points and hex.
 * Each block's `style` is element CSS only, which wins over `styles` and over the template.
 */
export type DocumentTree = {
  styles: ResolvedStyle;
  blocks: Block[];
  skipped: SkippedItem[];
};

export type ParseOptions = {
  styles?: StyleInput;
};

/** Lowest cascade layer, applied when `write` builds the .docx. */
export type DocxTemplate = {
  styles?: StyleInput;
  headings?: Partial<Record<HeadingLevel, StyleInput>>;
};

export type WriteOptions = {
  template?: DocxTemplate;
};

export type WriteResult = {
  blob: Blob;
  skipped: SkippedItem[];
};
