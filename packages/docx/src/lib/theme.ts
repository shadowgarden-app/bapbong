import { audit } from './audit.js';
import { attrOf, child, findDescendant, OoxmlNode } from './ooxml.js';

/** Resolve an OOXML `w:themeColor` (+ optional tint/shade hex) to "#RRGGBB". */
export type ThemeResolver = (
  themeColor: string,
  tint?: string,
  shade?: string,
) => string | undefined;

// clrScheme key → the w:themeColor names that reference it.
const ALIASES: Record<string, string[]> = {
  dk1: ['dark1', 'text1'],
  lt1: ['light1', 'background1'],
  dk2: ['dark2', 'text2'],
  lt2: ['light2', 'background2'],
  accent1: ['accent1'],
  accent2: ['accent2'],
  accent3: ['accent3'],
  accent4: ['accent4'],
  accent5: ['accent5'],
  accent6: ['accent6'],
  hlink: ['hyperlink'],
  folHlink: ['followedHyperlink'],
};

function schemeColor(el: OoxmlNode): string | undefined {
  const srgb = child(el, 'a:srgbClr');
  const srgbVal = attrOf(srgb, 'val');
  if (srgbVal) return `#${srgbVal.toUpperCase()}`;
  const sys = child(el, 'a:sysClr');
  // lastClr is the resolved system color; val ("window", "windowText") is
  // the symbolic name we'd otherwise map by hand. Both are read; lastClr wins.
  const lastClr = attrOf(sys, 'lastClr');
  const sysName = attrOf(sys, 'val');
  const sysVal = lastClr ?? (sysName ? SYS_COLORS[sysName] : undefined);
  if (sysVal) return `#${sysVal.toUpperCase().replace('#', '')}`;
  return undefined;
}

/** The two system colors OOXML themes actually use. */
const SYS_COLORS: Record<string, string> = {
  window: 'FFFFFF',
  windowText: '000000',
};

const clamp = (n: number) => Math.max(0, Math.min(255, Math.round(n)));
const hex2 = (n: number) =>
  clamp(n).toString(16).padStart(2, '0').toUpperCase();

/** Word's tint/shade as a simple linear RGB approximation (good enough for now). */
function applyTintShade(hex: string, tint?: string, shade?: string): string {
  let r = parseInt(hex.slice(1, 3), 16);
  let g = parseInt(hex.slice(3, 5), 16);
  let b = parseInt(hex.slice(5, 7), 16);
  if (shade !== undefined) {
    const f = parseInt(shade, 16) / 255;
    r *= f;
    g *= f;
    b *= f;
  }
  if (tint !== undefined) {
    const f = parseInt(tint, 16) / 255;
    r = r * f + 255 * (1 - f);
    g = g * f + 255 * (1 - f);
    b = b * f + 255 * (1 - f);
  }
  return `#${hex2(r)}${hex2(g)}${hex2(b)}`;
}

/** Resolve a `w:asciiTheme`-style slot token (majorHAnsi, minorEastAsia, …)
 *  to the theme's typeface name. */
export type ThemeFontResolver = (slot: string) => string | undefined;

/** Parse `a:fontScheme` (major/minor latin + eastAsia + cs typefaces). Word
 *  writes runs with ONLY `w:asciiTheme="minorHAnsi"` etc. — without this the
 *  run has no family at all and falls back to the app default font. */
export function buildThemeFontResolver(
  themeRoot: OoxmlNode | undefined,
): ThemeFontResolver {
  const scheme = findDescendant(themeRoot, 'a:fontScheme');
  const face = (
    group: OoxmlNode | undefined,
    tag: string,
  ): string | undefined => attrOf(child(group, tag), 'typeface') || undefined;

  const major = child(scheme, 'a:majorFont');
  const minor = child(scheme, 'a:minorFont');
  const map = new Map<string, string | undefined>([
    ['majorAscii', face(major, 'a:latin')],
    ['majorHAnsi', face(major, 'a:latin')],
    ['majorEastAsia', face(major, 'a:ea')],
    ['majorBidi', face(major, 'a:cs')],
    ['minorAscii', face(minor, 'a:latin')],
    ['minorHAnsi', face(minor, 'a:latin')],
    ['minorEastAsia', face(minor, 'a:ea')],
    ['minorBidi', face(minor, 'a:cs')],
  ]);
  // Script-specific alternates (a:font script="Hans" …) are deliberately not
  // itemized — one typeface per slot is what we render with.
  if (scheme) audit.markSubtree(scheme);

  return (slot) => map.get(slot);
}

export function buildThemeResolver(
  themeRoot: OoxmlNode | undefined,
): ThemeResolver {
  const map = new Map<string, string>();
  const scheme = findDescendant(themeRoot, 'a:clrScheme');
  for (const el of scheme?.children ?? []) {
    audit.mark(el);
    const key = el.name.startsWith('a:') ? el.name.slice(2) : el.name;
    const color = schemeColor(el);
    if (!color) continue;
    map.set(key, color);
    for (const alias of ALIASES[key] ?? []) map.set(alias, color);
  }

  return (themeColor, tint, shade) => {
    const base = map.get(themeColor);
    if (!base) return undefined;
    return tint === undefined && shade === undefined
      ? base
      : applyTintShade(base, tint, shade);
  };
}

// ── DrawingML colour ────────────────────────────────────────────────
// ST_ColorChoice is a CLOSED union of six elements, each able to carry a
// stack of transforms. Reading only two of the six (and none of the
// transforms) is how a shape ends up rendered in its base accent colour when
// Word shows "Accent 1, Lighter 40%" — that pair is written as schemeClr +
// lumMod + lumOff, and dropping the children keeps the val and loses the
// adjustment silently. Because the union is closed, covering it once closes
// it for good.

/** The subset of a:prstClr Word actually emits (shape defaults and the
 *  standard-colour row). Unlisted names resolve to undefined and fall through
 *  to whatever default the caller has. */
const PRESET_COLORS: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  cyan: '00FFFF',
  aqua: '00FFFF',
  magenta: 'FF00FF',
  fuchsia: 'FF00FF',
  gray: '808080',
  grey: '808080',
  darkGray: 'A9A9A9',
  lightGray: 'D3D3D3',
  silver: 'C0C0C0',
  maroon: '800000',
  olive: '808000',
  navy: '000080',
  purple: '800080',
  teal: '008080',
  lime: '00FF00',
  orange: 'FFA500',
  brown: 'A52A2A',
  pink: 'FFC0CB',
  gold: 'FFD700',
  indigo: '4B0082',
  violet: 'EE82EE',
};

const CHOICE_TAGS = [
  'a:srgbClr',
  'a:schemeClr',
  'a:sysClr',
  'a:prstClr',
  'a:scrgbClr',
  'a:hslClr',
];

/** OOXML percentages are thousandths of a percent: 40000 = 40%. */
function pct(el: OoxmlNode | undefined, attr: string): number | undefined {
  const v = attrOf(el, attr);
  if (v === undefined) return undefined;
  const n = Number(v.endsWith('%') ? v.slice(0, -1) : v) / 100000;
  return Number.isNaN(n) ? undefined : n;
}

/** Colour carried through a transform chain as floats in 0..1. Quantising to
 *  bytes between transforms is what put "Accent 1, Lighter 40%" two units off
 *  Word's swatch — lumMod and lumOff are written as a PAIR, so the
 *  intermediate value must never be rounded. */
type Rgb = [number, number, number];

const hexToRgb = (hex: string): Rgb => [
  parseInt(hex.slice(1, 3), 16) / 255,
  parseInt(hex.slice(3, 5), 16) / 255,
  parseInt(hex.slice(5, 7), 16) / 255,
];

const rgbToHex = ([r, g, b]: Rgb): string =>
  `#${hex2(r * 255)}${hex2(g * 255)}${hex2(b * 255)}`;

function rgbToHsl([r, g, b]: Rgb): [number, number, number] {
  const max = Math.max(r, g, b),
    min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h =
    max === r
      ? ((g - b) / d + (g < b ? 6 : 0)) / 6
      : max === g
        ? ((b - r) / d + 2) / 6
        : ((r - g) / d + 4) / 6;
  return [h, s, l];
}

function hslToRgb(h: number, s: number, l: number): Rgb {
  const sat = Math.max(0, Math.min(1, s));
  const lum = Math.max(0, Math.min(1, l));
  if (sat === 0) return [lum, lum, lum];
  const q = lum < 0.5 ? lum * (1 + sat) : lum + sat - lum * sat;
  const p = 2 * lum - q;
  const comp = (t: number) => {
    let x = t;
    if (x < 0) x += 1;
    if (x > 1) x -= 1;
    if (x < 1 / 6) return p + (q - p) * 6 * x;
    if (x < 1 / 2) return q;
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
    return p;
  };
  return [comp(h + 1 / 3), comp(h), comp(h - 1 / 3)];
}

/** Apply one colour transform. Luminance/saturation modifiers go through HSL
 *  — that is the pair Word writes for every "Lighter N%" swatch, and HSL
 *  reproduces its result exactly. a:alpha is deliberately absent: the colour
 *  model is opaque "#RRGGBB", so faking transparency against an unknown
 *  backdrop would be worse than leaving it visibly unhandled. */
function applyTransform(c: Rgb, el: OoxmlNode): Rgb {
  const f = pct(el, 'val');
  if (f === undefined) return c;
  const hsl = () => rgbToHsl(c);
  switch (el.name) {
    case 'a:lumMod': {
      const [h, s, l] = hsl();
      return hslToRgb(h, s, l * f);
    }
    case 'a:lumOff': {
      const [h, s, l] = hsl();
      return hslToRgb(h, s, l + f);
    }
    case 'a:satMod': {
      const [h, s, l] = hsl();
      return hslToRgb(h, s * f, l);
    }
    case 'a:shade':
      return [c[0] * f, c[1] * f, c[2] * f];
    case 'a:tint':
      return [c[0] * f + (1 - f), c[1] * f + (1 - f), c[2] * f + (1 - f)];
    default:
      return c;
  }
}

/** Resolve the ST_ColorChoice child of `parent` (an a:solidFill, a:lnRef,
 *  a:gs …) to "#RRGGBB", applying its transforms in document order.
 *
 *  `phClr` substitutes for `a:schemeClr val="phClr"`, the placeholder a theme's
 *  fmtScheme uses to mean "whichever colour the shape's style reference
 *  supplied". Without it every themed shape fill resolves to nothing. */
export function drawingColor(
  parent: OoxmlNode | undefined,
  resolveTheme: ThemeResolver,
  phClr?: string,
): string | undefined {
  if (!parent) return undefined;
  let el: OoxmlNode | undefined;
  for (const tag of CHOICE_TAGS) {
    el = child(parent, tag);
    if (el) break;
  }
  if (!el) return undefined;

  let base: string | undefined;
  switch (el.name) {
    case 'a:srgbClr': {
      const v = attrOf(el, 'val');
      base = v ? `#${v.toUpperCase()}` : undefined;
      break;
    }
    case 'a:schemeClr': {
      const v = attrOf(el, 'val');
      base = v === 'phClr' ? phClr : v ? resolveTheme(v) : undefined;
      break;
    }
    case 'a:sysClr': {
      // lastClr is Word's own resolution of the system colour; val is the
      // symbolic name we'd otherwise map by hand.
      const last = attrOf(el, 'lastClr');
      const name = attrOf(el, 'val');
      const v = last ?? (name ? SYS_COLORS[name] : undefined);
      base = v ? `#${v.toUpperCase().replace('#', '')}` : undefined;
      break;
    }
    case 'a:prstClr': {
      const v = attrOf(el, 'val');
      const hex = v ? PRESET_COLORS[v] : undefined;
      base = hex ? `#${hex}` : undefined;
      break;
    }
    case 'a:scrgbClr': {
      const r = pct(el, 'r'),
        g = pct(el, 'g'),
        b = pct(el, 'b');
      base =
        r !== undefined && g !== undefined && b !== undefined
          ? `#${hex2(r * 255)}${hex2(g * 255)}${hex2(b * 255)}`
          : undefined;
      break;
    }
    case 'a:hslClr': {
      const hueRaw = attrOf(el, 'hue');
      const h = hueRaw === undefined ? undefined : Number(hueRaw) / 21600000;
      const sat = pct(el, 'sat'),
        lum = pct(el, 'lum');
      base =
        h !== undefined && sat !== undefined && lum !== undefined
          ? rgbToHex(hslToRgb(h, sat, lum))
          : undefined;
      break;
    }
  }
  if (!base) return undefined;

  // Transforms compose in document order, so "lumMod 60% then lumOff 40%"
  // is not the same as the reverse.
  let out = hexToRgb(base);
  for (const t of el.children) {
    audit.mark(t);
    out = applyTransform(out, t);
  }
  return rgbToHex(out);
}

/**
 * Resolve an `a:lnRef` to the theme line style it points at.
 *
 * The twin of {@link ThemeFillResolver}, and the same indexing rule: `@idx`
 * is one-based into the format scheme's `a:lnStyleLst`. It hands back the
 * `a:ln` ELEMENT rather than parsed properties, because the caller already
 * knows how to read one — width, dash, cap and the rest come out of the same
 * code that reads a shape's own outline, so the two cannot drift apart.
 *
 * The colour is NOT resolved here: a shape's stroke colour already comes from
 * the ref itself (drawingColor over a:lnRef), which is the placeholder the
 * scheme entry would have substituted anyway.
 */
export type ThemeLineResolver = (
  lnRef: OoxmlNode | undefined,
) => OoxmlNode | undefined;

export function buildThemeLineResolver(
  themeRoot: OoxmlNode | undefined,
): ThemeLineResolver {
  const list = findDescendant(themeRoot, 'a:lnStyleLst');
  return (lnRef) => {
    if (!lnRef || !list) return undefined;
    const idx = Number(attrOf(lnRef, 'idx') ?? '0');
    if (!Number.isFinite(idx) || idx < 1) return undefined;
    // Indexed directly rather than through `children()`, which would mark the
    // whole list. Only the entry a document actually names is read, so the
    // ones nothing points at keep showing up in the coverage audit — which is
    // right: the day a file names entry 3 and we cannot paint what is in it,
    // that has to be visible.
    return list.children[idx - 1];
  };
}

/** Resolve an `a:fillRef` — a shape saying "fill me the way the theme's
 *  format scheme entry N does, using this colour as the placeholder". */
export type ThemeFillResolver = (
  fillRef: OoxmlNode | undefined,
) => string | undefined;

export function buildThemeFillResolver(
  themeRoot: OoxmlNode | undefined,
  resolveTheme: ThemeResolver,
): ThemeFillResolver {
  const list = findDescendant(themeRoot, 'a:fillStyleLst');
  const bgList = findDescendant(themeRoot, 'a:bgFillStyleLst');

  return (fillRef) => {
    if (!fillRef) return undefined;
    const idx = Number(attrOf(fillRef, 'idx') ?? '0');
    // The ref carries the placeholder colour the scheme entry fills in. Read
    // BEFORE the idx test, not after: returning early past it leaves the
    // child colour untouched, and the coverage audit cannot tell an untouched
    // node from an unsupported one.
    const phClr = drawingColor(fillRef, resolveTheme);
    if (!Number.isFinite(idx) || idx === 0) return undefined; // idx 0 = no fill
    const entry =
      idx >= 1000
        ? bgList?.children[idx - 1000]
        : (list?.children[idx - 1] ?? undefined);
    if (!entry) return phClr;
    if (entry.name === 'a:solidFill')
      return drawingColor(entry, resolveTheme, phClr) ?? phClr;
    // Gradient and pattern entries: we paint flat, so take the first stop —
    // closer than the bare placeholder, which ignores the scheme's tinting.
    const firstStop = child(child(entry, 'a:gsLst'), 'a:gs');
    return drawingColor(firstStop, resolveTheme, phClr) ?? phClr;
  };
}
