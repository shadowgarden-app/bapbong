/**
 * Colours as the document model keeps them: "#RRGGBB", upper case.
 *
 * Colours arrive in other shapes. A paste from a web page carries CSS —
 * browsers serialise every computed colour as `rgb(0, 0, 0)` — and a
 * document saved from such a paste held `<w:color w:val="rgb(0, 0, 0)"/>`,
 * which no consumer reads: our importer turned it into "#RGB(0, 0, 0)",
 * the canvas ignored that, and the text took whatever colour the previous
 * run had left (light blue, in the report). Every colour entering the
 * model, and every colour written to a file, goes through here.
 */

const NAMED: Record<string, string> = {
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
  silver: 'C0C0C0',
  maroon: '800000',
  olive: '808000',
  lime: '00FF00',
  navy: '000080',
  purple: '800080',
  teal: '008080',
  orange: 'FFA500',
};

const byte = (s: string): number | null => {
  const t = s.trim();
  if (t.endsWith('%')) {
    const p = Number(t.slice(0, -1));
    return Number.isFinite(p)
      ? Math.round(Math.min(100, Math.max(0, p)) * 2.55)
      : null;
  }
  const n = Number(t);
  return Number.isFinite(n) ? Math.round(Math.min(255, Math.max(0, n))) : null;
};

const hex2 = (n: number) => n.toString(16).padStart(2, '0').toUpperCase();

/**
 * "#RRGGBB" for a colour written as hex (with or without "#", 3 or 6
 * digits), `rgb()` / `rgba()` (commas or spaces), or a basic CSS name.
 * Undefined for anything else — `auto`, `transparent`, a fully transparent
 * rgba, `currentcolor`, a theme name: no colour, so the default applies.
 */
export function hexColor(value: string | null | undefined): string | undefined {
  if (!value) return undefined;
  const v = value.trim().toLowerCase();
  const hex = /^#?([0-9a-f]{6}|[0-9a-f]{3})$/.exec(v);
  if (hex) {
    const h =
      hex[1].length === 3 ? [...hex[1]].map((c) => c + c).join('') : hex[1];
    return `#${h.toUpperCase()}`;
  }
  const fn = /^rgba?\(([^)]*)\)$/.exec(v);
  if (fn) {
    const parts = fn[1].split(/\s*[,/]\s*|\s+/).filter(Boolean);
    if (parts.length < 3) return undefined;
    const alpha = parts[3];
    if (alpha !== undefined) {
      const a = alpha.endsWith('%')
        ? Number(alpha.slice(0, -1)) / 100
        : Number(alpha);
      if (Number.isFinite(a) && a === 0) return undefined;
    }
    const rgb = parts.slice(0, 3).map(byte);
    if (rgb.some((c) => c === null)) return undefined;
    return `#${(rgb as number[]).map(hex2).join('')}`;
  }
  const named = NAMED[v];
  return named ? `#${named}` : undefined;
}
