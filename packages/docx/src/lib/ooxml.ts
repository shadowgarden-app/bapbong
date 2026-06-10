import { XMLParser } from 'fast-xml-parser';

/** Shared OOXML parser. `trimValues:false` keeps significant whitespace in
 *  <w:t> (e.g. xml:space="preserve" "Hello "). Repeated elements are forced to
 *  arrays so walkers can treat single/multiple uniformly. */
export const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false,
  isArray: (name) =>
    name === 'w:p' || name === 'w:r' || name === 'w:t' || name === 'w:style',
});

export function toArray<T>(v: T | T[] | undefined | null): T[] {
  if (v == null) return [];
  return Array.isArray(v) ? v : [v];
}

export function asRecord(v: unknown): Record<string, unknown> | undefined {
  return v && typeof v === 'object' ? (v as Record<string, unknown>) : undefined;
}

/** Read an attribute (e.g. `@_w:val`) off an element, coercing to string. */
export function attr(el: unknown, name: string): string | undefined {
  const v = asRecord(el)?.[name];
  return v == null ? undefined : String(v);
}

const OFF = new Set(['false', '0', 'off']);

/** OOXML on/off toggle. Returns `undefined` when the element is absent (so the
 *  cascade leaves the inherited value untouched), else the resolved boolean. */
export function toggle(el: unknown): boolean | undefined {
  if (el === undefined) return undefined;
  const v = attr(el, '@_w:val');
  return v === undefined ? true : !OFF.has(v.toLowerCase());
}

/** Underline variant: `w:val="none"` (and falsy) means off. */
export function underlineToggle(el: unknown): boolean | undefined {
  if (el === undefined) return undefined;
  const v = attr(el, '@_w:val');
  if (v === undefined) return true;
  const lower = v.toLowerCase();
  return lower !== 'none' && !OFF.has(lower);
}

/** Normalized run (character) properties, independent of the OOXML shape. */
export interface RunProps {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string; // "#RRGGBB"
  sizePt?: number; // points
  fontFamily?: string;
}

/** Parse a `w:rPr` element into normalized run properties. Only sets a key
 *  when the property is present, so merging preserves inherited values. */
export function parseRunProps(rPr: Record<string, unknown> | undefined): RunProps {
  if (!rPr) return {};
  const props: RunProps = {};

  const b = toggle(rPr['w:b']);
  if (b !== undefined) props.bold = b;
  const i = toggle(rPr['w:i']);
  if (i !== undefined) props.italic = i;
  const u = underlineToggle(rPr['w:u']);
  if (u !== undefined) props.underline = u;
  const s = toggle(rPr['w:strike']);
  if (s !== undefined) props.strike = s;

  const color = attr(rPr['w:color'], '@_w:val');
  if (color && color.toLowerCase() !== 'auto') {
    props.color = color.startsWith('#') ? color.toUpperCase() : `#${color.toUpperCase()}`;
  }

  const sz = attr(rPr['w:sz'], '@_w:val');
  if (sz !== undefined) {
    const halfPoints = Number(sz);
    if (!Number.isNaN(halfPoints)) props.sizePt = halfPoints / 2;
  }

  const family = attr(rPr['w:rFonts'], '@_w:ascii') ?? attr(rPr['w:rFonts'], '@_w:hAnsi');
  if (family) props.fontFamily = family;

  return props;
}

/** Merge run properties; defined keys in `over` win over `base`. */
export function mergeRunProps(base: RunProps, over: RunProps): RunProps {
  return { ...base, ...over };
}
