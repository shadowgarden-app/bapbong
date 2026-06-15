import { XMLParser } from 'fast-xml-parser';

/**
 * A normalized OOXML element. We parse with `preserveOrder: true` (so the
 * document's block order — paragraphs vs tables — is faithful) and flatten the
 * verbose fast-xml-parser shape into this simple tree: ordered `children`,
 * attribute map (without the `@_` prefix), and concatenated `text`.
 */
export interface OoxmlNode {
  name: string;
  attrs: Record<string, string>;
  children: OoxmlNode[];
  text: string;
}

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  // Keep significant whitespace in <w:t> (e.g. xml:space="preserve" "Hello ").
  trimValues: false,
});

function buildNode(entry: Record<string, unknown>): OoxmlNode {
  const name = Object.keys(entry).find((k) => k !== ':@') ?? '';

  const attrs: Record<string, string> = {};
  const rawAttrs = entry[':@'] as Record<string, unknown> | undefined;
  if (rawAttrs) {
    for (const [k, v] of Object.entries(rawAttrs)) {
      attrs[k.startsWith('@_') ? k.slice(2) : k] = v == null ? '' : String(v);
    }
  }

  const raw = entry[name];
  const rawChildren = Array.isArray(raw) ? (raw as Record<string, unknown>[]) : [];
  const children: OoxmlNode[] = [];
  let text = '';
  for (const c of rawChildren) {
    if (Object.prototype.hasOwnProperty.call(c, '#text')) {
      text += String((c as Record<string, unknown>)['#text'] ?? '');
    } else {
      children.push(buildNode(c));
    }
  }
  return { name, attrs, children, text };
}

/** Parse an OOXML part into a synthetic `#root` node holding the top-level elements. */
export function parseXml(xml: string): OoxmlNode {
  const top = parser.parse(xml) as Record<string, unknown>[];
  const children = top
    .filter((e) => {
      const key = Object.keys(e).find((k) => k !== ':@');
      return !!key && key !== '#text' && !key.startsWith('?');
    })
    .map(buildNode);
  return { name: '#root', attrs: {}, children, text: '' };
}

/** First child element with the given tag name. */
export function child(node: OoxmlNode | undefined, name: string): OoxmlNode | undefined {
  return node?.children.find((c) => c.name === name);
}

/** All child elements with the given tag name (in document order). */
export function children(node: OoxmlNode | undefined, name: string): OoxmlNode[] {
  return node ? node.children.filter((c) => c.name === name) : [];
}

/** An attribute value (name without the `@_`/`w:` mangling, e.g. "w:val"). */
export function attrOf(node: OoxmlNode | undefined, name: string): string | undefined {
  return node?.attrs[name];
}

/** Depth-first search for the first descendant with the given tag name. */
export function findDescendant(node: OoxmlNode | undefined, name: string): OoxmlNode | undefined {
  if (!node) return undefined;
  for (const c of node.children) {
    if (c.name === name) return c;
    const found = findDescendant(c, name);
    if (found) return found;
  }
  return undefined;
}

const OFF = new Set(['false', '0', 'off']);

/** OOXML on/off toggle. `undefined` when absent (cascade keeps inherited value). */
function toggle(el: OoxmlNode | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = attrOf(el, 'w:val');
  return v === undefined ? true : !OFF.has(v.toLowerCase());
}

function underlineToggle(el: OoxmlNode | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = attrOf(el, 'w:val');
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
  highlight?: string; // background "#RRGGBB" (w:highlight / w:shd w:fill)
}

/** Word's 16 named highlight colors → hex. */
const HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF',
  blue: '#0000FF', red: '#FF0000', darkBlue: '#000080', darkCyan: '#008080',
  darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0',
  black: '#000000', white: '#FFFFFF',
};

/** Normalize an OOXML hex color ("FF0000" or "#ff0000") to "#FF0000". */
export function normalizeHex(v: string | undefined): string | undefined {
  if (!v || v.toLowerCase() === 'auto') return undefined;
  return v.startsWith('#') ? v.toUpperCase() : `#${v.toUpperCase()}`;
}

/** Resolve a `w:themeColor` (+ optional tint/shade) to "#RRGGBB". */
export type ThemeColorResolver = (
  themeColor: string,
  tint?: string,
  shade?: string,
) => string | undefined;

/** Parse a `w:rPr` element into normalized run properties (only present keys). */
export function parseRunProps(
  rPr: OoxmlNode | undefined,
  resolveTheme?: ThemeColorResolver,
): RunProps {
  if (!rPr) return {};
  const props: RunProps = {};

  const b = toggle(child(rPr, 'w:b'));
  if (b !== undefined) props.bold = b;
  const i = toggle(child(rPr, 'w:i'));
  if (i !== undefined) props.italic = i;
  const u = underlineToggle(child(rPr, 'w:u'));
  if (u !== undefined) props.underline = u;
  const s = toggle(child(rPr, 'w:strike'));
  if (s !== undefined) props.strike = s;

  const colorEl = child(rPr, 'w:color');
  const colorVal = attrOf(colorEl, 'w:val');
  if (colorVal && colorVal.toLowerCase() !== 'auto') {
    props.color = colorVal.startsWith('#') ? colorVal.toUpperCase() : `#${colorVal.toUpperCase()}`;
  } else if (resolveTheme) {
    const themeColor = attrOf(colorEl, 'w:themeColor');
    if (themeColor) {
      const hex = resolveTheme(
        themeColor,
        attrOf(colorEl, 'w:themeTint'),
        attrOf(colorEl, 'w:themeShade'),
      );
      if (hex) props.color = hex;
    }
  }

  const sz = attrOf(child(rPr, 'w:sz'), 'w:val');
  if (sz !== undefined) {
    const halfPoints = Number(sz);
    if (!Number.isNaN(halfPoints)) props.sizePt = halfPoints / 2;
  }

  const rFonts = child(rPr, 'w:rFonts');
  const family = attrOf(rFonts, 'w:ascii') ?? attrOf(rFonts, 'w:hAnsi');
  if (family) props.fontFamily = family;

  // Background: w:highlight (named) takes precedence, else w:shd w:fill (hex).
  const hl = attrOf(child(rPr, 'w:highlight'), 'w:val');
  if (hl && hl !== 'none') {
    props.highlight = HIGHLIGHT_COLORS[hl] ?? normalizeHex(hl);
  } else {
    const hex = normalizeHex(attrOf(child(rPr, 'w:shd'), 'w:fill'));
    if (hex) props.highlight = hex;
  }

  return props;
}

/** Merge run properties; defined keys in `over` win over `base`. */
export function mergeRunProps(base: RunProps, over: RunProps): RunProps {
  return { ...base, ...over };
}
