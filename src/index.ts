import { write } from './write-docx.js';
import { parse } from './parse-html.js';

export { PagekitError } from './errors.js';
export { parse } from './parse-html.js';
export { write } from './write-docx.js';
export type {
  Alignment,
  Block,
  DocxTemplate,
  DocumentTree,
  HeadingLevel,
  HyperlinkInline,
  InlineNode,
  LineBreakInline,
  ListBlock,
  ListItem,
  ParagraphBlock,
  ParseOptions,
  ResolvedStyle,
  SkippedItem,
  StyleInput,
  TableBlock,
  TableCell,
  TableRow,
  TextInline,
  WriteOptions,
  WriteResult,
} from './types.js';

/** Browser-only HTML to .docx. */
export const pagekit = { parse, write };
