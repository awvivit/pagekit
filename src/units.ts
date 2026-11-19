import type { Alignment, ResolvedStyle, StyleInput } from './types.js';

const STYLE_KEYS: (keyof ResolvedStyle)[] = [
  'fontFamily',
  'fontSizePt',
  'color',
  'bold',
  'italic',
  'underline',
  'alignment',
  'spaceBeforePt',
  'spaceAfterPt',
  'lineSpacingPt',
  'indentLeftPt',
];

/** Later layers replace earlier ones. `false` overrides `true`. */
export function mergeStyles(...layers: Array<ResolvedStyle | undefined>): ResolvedStyle {
  const merged: ResolvedStyle = {};
  for (const layer of layers) {
    if (!layer) continue;
    for (const key of STYLE_KEYS) {
      const value = layer[key];
      if (value !== undefined) {
        Object.assign(merged, { [key]: value });
      }
    }
  }
  return merged;
}

export function characterStyleOf(style: ResolvedStyle | undefined): ResolvedStyle {
  if (!style) return {};
  return {
    fontFamily: style.fontFamily,
    fontSizePt: style.fontSizePt,
    color: style.color,
    bold: style.bold,
    italic: style.italic,
    underline: style.underline,
  };
}

export function paragraphStyleOf(style: ResolvedStyle | undefined): ResolvedStyle {
  if (!style) return {};
  return {
    alignment: style.alignment,
    spaceBeforePt: style.spaceBeforePt,
    spaceAfterPt: style.spaceAfterPt,
    lineSpacingPt: style.lineSpacingPt,
    indentLeftPt: style.indentLeftPt,
  };
}

export function roundPt(value: number): number {
  return Math.round(value * 100) / 100;
}

/** Convert a CSS length to points. Unitless strings are points (the Pagekit API). */
export function cssLengthToPt(input: string, basePt = 12): number | undefined {
  const value = input.trim().toLowerCase();
  if (!value || value === 'auto' || value === 'normal' || value === 'inherit' || value === 'initial') {
    return undefined;
  }
  const match = /^(-?[0-9]*\.?[0-9]+)(px|pt|pc|in|cm|mm|em|rem|%)?$/.exec(value);
  if (!match) return undefined;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount)) return undefined;
  const unit = match[2] ?? 'pt';
  switch (unit) {
    case 'pt':
      return roundPt(amount);
    case 'px':
      return roundPt(amount * 0.75);
    case 'pc':
      return roundPt(amount * 12);
    case 'in':
      return roundPt(amount * 72);
    case 'cm':
      return roundPt(amount * 28.3464567);
    case 'mm':
      return roundPt(amount * 2.83464567);
    case 'em':
    case 'rem':
      return roundPt(amount * basePt);
    case '%':
      return roundPt((amount / 100) * basePt);
    default:
      return undefined;
  }
}

export function firstFontFamily(value: string): string | undefined {
  const first = value.split(',')[0]?.trim().replace(/^['"]+|['"]+$/g, '');
  if (!first) return undefined;
  const lower = first.toLowerCase();
  if (lower === 'inherit' || lower === 'initial' || lower === 'unset') return undefined;
  return first;
}

function channel(value: string): number {
  if (value.endsWith('%')) return Math.round((Number(value.slice(0, -1)) / 100) * 255);
  return Number(value);
}

function hexByte(value: number): string {
  return Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, '0');
}

function composite(red: number, green: number, blue: number, alpha: number): string | undefined {
  if (!Number.isFinite(alpha) || alpha <= 0) return undefined;
  const cover = Math.min(1, alpha);
  const blend = (channelValue: number) => channelValue * cover + 255 * (1 - cover);
  return `#${hexByte(blend(red))}${hexByte(blend(green))}${hexByte(blend(blue))}`.toUpperCase();
}

/** Convert a CSS color to `#RRGGBB`, compositing transparency over white. */
export function cssColorToHex(input: string): string | undefined {
  const value = input.trim().toLowerCase();
  if (!value || value === 'transparent' || value === 'inherit' || value === 'initial' || value === 'unset' || value === 'currentcolor') {
    return undefined;
  }
  const hex = /^#([0-9a-f]{3,8})$/.exec(value);
  if (hex) {
    const raw = hex[1] ?? '';
    if (raw.length === 3 || raw.length === 4) {
      const red = Number.parseInt((raw[0] ?? '') + (raw[0] ?? ''), 16);
      const green = Number.parseInt((raw[1] ?? '') + (raw[1] ?? ''), 16);
      const blue = Number.parseInt((raw[2] ?? '') + (raw[2] ?? ''), 16);
      const alpha = raw.length === 4 ? Number.parseInt((raw[3] ?? '') + (raw[3] ?? ''), 16) / 255 : 1;
      return composite(red, green, blue, alpha);
    }
    if (raw.length === 6 || raw.length === 8) {
      const red = Number.parseInt(raw.slice(0, 2), 16);
      const green = Number.parseInt(raw.slice(2, 4), 16);
      const blue = Number.parseInt(raw.slice(4, 6), 16);
      const alpha = raw.length === 8 ? Number.parseInt(raw.slice(6, 8), 16) / 255 : 1;
      return composite(red, green, blue, alpha);
    }
  }
  const rgb = /^rgba?\(\s*([0-9.]+%?)\s*[, ]\s*([0-9.]+%?)\s*[, ]\s*([0-9.]+%?)(?:\s*[,/]\s*([0-9.]+%?))?\s*\)$/.exec(value);
  if (rgb) {
    const alphaRaw = rgb[4];
    let alpha = 1;
    if (alphaRaw) {
      alpha = alphaRaw.endsWith('%') ? Number(alphaRaw.slice(0, -1)) / 100 : Number(alphaRaw);
    }
    return composite(channel(rgb[1] ?? '0'), channel(rgb[2] ?? '0'), channel(rgb[3] ?? '0'), alpha);
  }
  return browserNamedColor(value);
}

function browserNamedColor(value: string): string | undefined {
  if (typeof document === 'undefined') return undefined;
  const probe = document.createElement('span');
  probe.style.color = value;
  if (!probe.style.color) return undefined;
  const parent = document.body ?? document.documentElement;
  parent.appendChild(probe);
  const computed = probe.ownerDocument.defaultView?.getComputedStyle(probe).color ?? '';
  probe.remove();
  if (!computed || computed.trim().toLowerCase() === value) return undefined;
  return cssColorToHex(computed);
}

export function normalizeAlignment(value: string | undefined): Alignment | undefined {
  switch (value) {
    case 'left':
    case 'start':
      return 'left';
    case 'right':
    case 'end':
      return 'right';
    case 'center':
      return 'center';
    case 'justify':
    case 'both':
      return 'justify';
    default:
      return undefined;
  }
}

function lengthField(value: number | string | undefined, basePt?: number): number | undefined {
  if (typeof value === 'number') {
    if (!Number.isFinite(value) || value < 0) return undefined;
    return roundPt(value);
  }
  if (typeof value === 'string') return cssLengthToPt(value, basePt);
  return undefined;
}

export function normalizeStyleInput(input: StyleInput | undefined): ResolvedStyle {
  if (!input) return {};
  const style: ResolvedStyle = {};
  if (input.fontFamily) {
    const family = firstFontFamily(input.fontFamily);
    if (family) style.fontFamily = family;
  }
  const fontSize = lengthField(input.fontSize);
  if (fontSize !== undefined && fontSize > 0) style.fontSizePt = fontSize;
  if (input.color) {
    const color = cssColorToHex(input.color);
    if (color) style.color = color;
  }
  if (input.bold !== undefined) style.bold = input.bold;
  if (input.italic !== undefined) style.italic = input.italic;
  if (input.underline !== undefined) style.underline = input.underline;
  const alignment = normalizeAlignment(input.alignment);
  if (alignment) style.alignment = alignment;
  const before = lengthField(input.spaceBefore);
  if (before !== undefined) style.spaceBeforePt = before;
  const after = lengthField(input.spaceAfter);
  if (after !== undefined) style.spaceAfterPt = after;
  const line = lengthField(input.lineSpacing, fontSize);
  if (line !== undefined && line > 0) style.lineSpacingPt = line;
  const indent = lengthField(input.indentLeft);
  if (indent !== undefined) style.indentLeftPt = indent;
  return style;
}

/** Drop characters XML 1.0 forbids in text nodes. */
export function stripIllegalXml(text: string): string {
  let result = '';
  for (const char of text) {
    const code = char.charCodeAt(0);
    const illegal = code <= 0x08 || code === 0x0b || code === 0x0c || (code >= 0x0e && code <= 0x1f);
    if (!illegal) result += char;
  }
  return result;
}

export function twips(points: number): string {
  return String(Math.max(0, Math.round(points * 20)));
}

export function halfPoints(points: number): string {
  return String(Math.max(1, Math.round(points * 2)));
}
