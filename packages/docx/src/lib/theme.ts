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
  const sysVal = attrOf(sys, 'lastClr') ?? attrOf(sys, 'val');
  if (sysVal) return `#${sysVal.toUpperCase()}`;
  return undefined;
}

const clamp = (n: number) => Math.max(0, Math.min(255, Math.round(n)));
const hex2 = (n: number) => clamp(n).toString(16).padStart(2, '0').toUpperCase();

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

export function buildThemeResolver(themeRoot: OoxmlNode | undefined): ThemeResolver {
  const map = new Map<string, string>();
  const scheme = findDescendant(themeRoot, 'a:clrScheme');
  for (const el of scheme?.children ?? []) {
    const key = el.name.startsWith('a:') ? el.name.slice(2) : el.name;
    const color = schemeColor(el);
    if (!color) continue;
    map.set(key, color);
    for (const alias of ALIASES[key] ?? []) map.set(alias, color);
  }

  return (themeColor, tint, shade) => {
    const base = map.get(themeColor);
    if (!base) return undefined;
    return tint === undefined && shade === undefined ? base : applyTintShade(base, tint, shade);
  };
}
