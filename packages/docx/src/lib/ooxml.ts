import { XMLParser } from 'fast-xml-parser';
import { audit } from './audit.js';

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
  // Text is text — never strnum it. The default (true) mangles document
  // content: "100.000" → 100, "00" → 0, "1." → 1 — Word splits numbers like
  // "1.500.000" across runs (rsid), and each fragment got destroyed.
  parseTagValue: false,
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
  const rawChildren = Array.isArray(raw)
    ? (raw as Record<string, unknown>[])
    : [];
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

// The four accessors below are the importer's only doorway into the parsed
// tree, so they double as the XML-audit's coverage probes (see audit.ts) —
// one boolean check each when the audit flag is off.

const escText = (s: string): string =>
  s.replace(
    /[&<>]/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c] as string,
  );
const escAttrVal = (s: string): string =>
  s.replace(
    /[&<>"]/g,
    (c) =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c] as string,
  );

/** Serialize a parsed node back to XML (carry-through fidelity). `keep`
 *  filters elements AND attribute names — carried fragments are re-embedded
 *  into a generated document whose root declares only OUR namespaces, so
 *  callers restrict to prefixes that stay well-formed there. Mixed content
 *  order is not preserved (children first, then text) — the OOXML property
 *  bags this serves have no mixed content. */
export function serializeOoxml(
  node: OoxmlNode,
  keep: { element(name: string): boolean; attr(name: string): boolean },
): string {
  if (!keep.element(node.name)) return '';
  const attrs = Object.entries(node.attrs)
    .filter(([k]) => keep.attr(k))
    .map(([k, v]) => ` ${k}="${escAttrVal(v)}"`)
    .join('');
  const inner =
    node.children.map((c) => serializeOoxml(c, keep)).join('') +
    (node.text ? escText(node.text) : '');
  return inner
    ? `<${node.name}${attrs}>${inner}</${node.name}>`
    : `<${node.name}${attrs}/>`;
}

/** First child element with the given tag name. */
export function child(
  node: OoxmlNode | undefined,
  name: string,
): OoxmlNode | undefined {
  const found = node?.children.find((c) => c.name === name);
  audit.mark(found);
  return found;
}

/** All child elements with the given tag name (in document order). */
export function children(
  node: OoxmlNode | undefined,
  name: string,
): OoxmlNode[] {
  const out = node ? node.children.filter((c) => c.name === name) : [];
  audit.markAll(out);
  return out;
}

/** An attribute value (name without the `@_`/`w:` mangling, e.g. "w:val"). */
export function attrOf(
  node: OoxmlNode | undefined,
  name: string,
): string | undefined {
  audit.markAttr(node, name);
  return node?.attrs[name];
}

/** Depth-first search for the first descendant with the given tag name.
 *  Audit: the found node AND the container chain leading to it are marked —
 *  the containers were structurally traversed to reach a consumed node. */
export function findDescendant(
  node: OoxmlNode | undefined,
  name: string,
): OoxmlNode | undefined {
  if (!node) return undefined;
  for (const c of node.children) {
    if (c.name === name) {
      audit.mark(c);
      return c;
    }
    const found = findDescendant(c, name);
    if (found) {
      audit.mark(c);
      return found;
    }
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
  dstrike?: boolean; // w:dstrike — double strikethrough
  smallCaps?: boolean; // w:smallCaps
  color?: string; // "#RRGGBB"
  sizePt?: number; // points
  fontFamily?: string;
  highlight?: string; // background "#RRGGBB" (w:highlight / w:shd w:fill)
  vertAlign?: 'super' | 'sub'; // w:vertAlign
  /** w:position — baseline shift in HALF-POINTS, positive up. Kept in the
   *  document's own unit; the layout converts once, at the point of use. */
  position?: number;
}

/** Word's 16 named highlight colors → hex. */
const HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: '#FFFF00',
  green: '#00FF00',
  cyan: '#00FFFF',
  magenta: '#FF00FF',
  blue: '#0000FF',
  red: '#FF0000',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#C0C0C0',
  black: '#000000',
  white: '#FFFFFF',
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

/** Resolve a `w:asciiTheme`-style slot token to the theme's typeface. */
export type ThemeFontResolver = (slot: string) => string | undefined;

/** Effective background of a `w:shd`. A `solid` pattern paints the PATTERN
 *  color (w:color) at 100%, everything else paints the fill — literal
 *  `w:fill` hex first, then the theme fill. Pattern percentages (pct25 …)
 *  are not blended; their fill is the closest paint we do. */
export function shdFill(
  shd: OoxmlNode | undefined,
  resolveTheme?: ThemeColorResolver,
): string | undefined {
  if (!shd) return undefined;
  const val = attrOf(shd, 'w:val');
  const color = attrOf(shd, 'w:color');
  const fill = normalizeHex(attrOf(shd, 'w:fill'));
  const themeFill = attrOf(shd, 'w:themeFill');
  const themeFillTint = attrOf(shd, 'w:themeFillTint');
  const themeFillShade = attrOf(shd, 'w:themeFillShade');
  if (val === 'solid') {
    const solid = normalizeHex(color);
    if (solid) return solid;
  }
  return (
    fill ??
    (themeFill && resolveTheme
      ? resolveTheme(themeFill, themeFillTint, themeFillShade)
      : undefined)
  );
}

/** Parse a `w:rPr` element into normalized run properties (only present keys). */
export function parseRunProps(
  rPr: OoxmlNode | undefined,
  resolveTheme?: ThemeColorResolver,
  resolveFont?: ThemeFontResolver,
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
  const ds = toggle(child(rPr, 'w:dstrike'));
  if (ds !== undefined) props.dstrike = ds;
  const sc = toggle(child(rPr, 'w:smallCaps'));
  if (sc !== undefined) props.smallCaps = sc;

  // Color: the literal w:val is Word's own baked rendering of the theme
  // reference, so it wins (no tint/shade approximation error); the theme
  // attrs are the fallback when only the reference was written. All four
  // attrs are read up front — each is part of this resolution.
  const colorEl = child(rPr, 'w:color');
  const colorVal = attrOf(colorEl, 'w:val');
  const themeColor = attrOf(colorEl, 'w:themeColor');
  const themeTint = attrOf(colorEl, 'w:themeTint');
  const themeShade = attrOf(colorEl, 'w:themeShade');
  if (colorVal && colorVal.toLowerCase() !== 'auto') {
    props.color = colorVal.startsWith('#')
      ? colorVal.toUpperCase()
      : `#${colorVal.toUpperCase()}`;
  } else if (resolveTheme && themeColor) {
    const hex = resolveTheme(themeColor, themeTint, themeShade);
    if (hex) props.color = hex;
  }

  const sz = attrOf(child(rPr, 'w:sz'), 'w:val');
  if (sz !== undefined) {
    const halfPoints = Number(sz);
    if (!Number.isNaN(halfPoints)) props.sizePt = halfPoints / 2;
  }

  // Font family: literal attrs first (ascii → hAnsi → eastAsia → cs), then
  // their theme-slot twins via the theme's fontScheme — Word often writes
  // ONLY `w:asciiTheme="minorHAnsi"` and no literal name at all. All eight
  // attrs are read up front (not short-circuited) — each is genuinely part
  // of the fallback chain.
  const rFonts = child(rPr, 'w:rFonts');
  const ascii = attrOf(rFonts, 'w:ascii');
  const hAnsi = attrOf(rFonts, 'w:hAnsi');
  const eastAsia = attrOf(rFonts, 'w:eastAsia');
  const cs = attrOf(rFonts, 'w:cs');
  const asciiTheme = attrOf(rFonts, 'w:asciiTheme');
  const hAnsiTheme = attrOf(rFonts, 'w:hAnsiTheme');
  const eastAsiaTheme = attrOf(rFonts, 'w:eastAsiaTheme');
  const cstheme = attrOf(rFonts, 'w:cstheme');
  const themed = (slot: string | undefined): string | undefined =>
    slot !== undefined && resolveFont ? resolveFont(slot) : undefined;
  const family =
    ascii ??
    hAnsi ??
    themed(asciiTheme) ??
    themed(hAnsiTheme) ??
    eastAsia ??
    themed(eastAsiaTheme) ??
    cs ??
    themed(cstheme);
  if (family) props.fontFamily = family;

  const va = attrOf(child(rPr, 'w:vertAlign'), 'w:val');
  if (va === 'superscript') props.vertAlign = 'super';
  else if (va === 'subscript') props.vertAlign = 'sub';

  // w:position is independent of vertAlign: it moves the baseline without
  // resizing the glyphs, and a run may carry both.
  const pos = attrOf(child(rPr, 'w:position'), 'w:val');
  if (pos !== undefined) {
    const hp = Number(pos);
    if (!Number.isNaN(hp)) props.position = hp;
  }

  // Background: w:highlight (named) takes precedence, else w:shd (solid
  // pattern color / fill / theme fill).
  const hl = attrOf(child(rPr, 'w:highlight'), 'w:val');
  if (hl && hl !== 'none') {
    props.highlight = HIGHLIGHT_COLORS[hl] ?? normalizeHex(hl);
  } else {
    const hex = shdFill(child(rPr, 'w:shd'), resolveTheme);
    if (hex) props.highlight = hex;
  }

  return props;
}

/** Merge run properties; defined keys in `over` win over `base`. */
export function mergeRunProps(base: RunProps, over: RunProps): RunProps {
  return { ...base, ...over };
}
