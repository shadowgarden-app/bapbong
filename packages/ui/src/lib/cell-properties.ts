import type { BorderStyle } from '@shadow-garden/bapbong-contracts';
import { colorButton } from './color-picker.js';
import { Dialog } from './dialog.js';
import { injectStyle } from './internal.js';

/** Which edges a border preset paints (Outside/Inside are position-dependent
 *  across a multi-cell block — the host resolves them per cell). */
export type BorderPreset =
  | 'all'
  | 'none'
  | 'outside'
  | 'inside'
  | 'top'
  | 'bottom'
  | 'left'
  | 'right'
  | 'insideH'
  | 'insideV';

/** The chosen border pen + which preset to apply it with. */
export interface BorderChoice {
  preset: BorderPreset;
  /** Stroke width in CSS px (converted from the pt picker). */
  width: number;
  style: BorderStyle;
  color: string;
}

export interface CellPropsResult {
  background: string | null;
  vAlign: 'top' | 'center' | 'bottom';
  /** Present only if the user picked a border preset (else borders untouched). */
  border?: BorderChoice;
}

export interface CellPropertiesOptions {
  initial: { background: string | null; vAlign: 'top' | 'center' | 'bottom' };
  /** Disable Inside / Inside-H / Inside-V (meaningless for a 1×1 selection). */
  singleCell: boolean;
  onApply: (result: CellPropsResult) => void;
}

const PT_WIDTHS = [0.5, 0.75, 1, 1.5, 2.25, 3];
const STYLES: ReadonlyArray<{ key: BorderStyle; label: string }> = [
  { key: 'solid', label: 'Solid' },
  { key: 'dashed', label: 'Dashed' },
  { key: 'dotted', label: 'Dotted' },
  { key: 'double', label: 'Double' },
];
const VALIGNS: ReadonlyArray<{
  key: CellPropsResult['vAlign'];
  label: string;
}> = [
  { key: 'top', label: 'Top' },
  { key: 'center', label: 'Middle' },
  { key: 'bottom', label: 'Bottom' },
];
const PRESETS: ReadonlyArray<{
  key: BorderPreset;
  label: string;
  sides: number[];
  inside: boolean;
}> = [
  {
    key: 'all',
    label: 'All borders',
    sides: [1, 1, 1, 1, 1, 1],
    inside: false,
  },
  { key: 'none', label: 'No border', sides: [0, 0, 0, 0, 0, 0], inside: false },
  {
    key: 'outside',
    label: 'Outside',
    sides: [1, 1, 1, 1, 0, 0],
    inside: false,
  },
  { key: 'inside', label: 'Inside', sides: [0, 0, 0, 0, 1, 1], inside: true },
  { key: 'top', label: 'Top', sides: [1, 0, 0, 0, 0, 0], inside: false },
  { key: 'bottom', label: 'Bottom', sides: [0, 0, 1, 0, 0, 0], inside: false },
  { key: 'left', label: 'Left', sides: [0, 0, 0, 1, 0, 0], inside: false },
  { key: 'right', label: 'Right', sides: [0, 1, 0, 0, 0, 0], inside: false },
  {
    key: 'insideH',
    label: 'Inside horizontal',
    sides: [0, 0, 0, 0, 1, 0],
    inside: true,
  },
  {
    key: 'insideV',
    label: 'Inside vertical',
    sides: [0, 0, 0, 0, 0, 1],
    inside: true,
  },
];

const STYLE = `
.bb-cp{display:flex;flex-direction:column;gap:16px;min-width:320px;color:var(--bb-ui-fg,#2c2c2a)}
.bb-cp *{box-sizing:border-box}
.bb-cp-label{font-size:12px;opacity:.7;margin-bottom:8px}
.bb-cp-seg{display:inline-flex;border:0.5px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;overflow:hidden}
.bb-cp-segbtn{height:32px;padding:0 16px;border:0;border-right:0.5px solid var(--bb-ui-border,#e3e3e0);background:transparent;color:inherit;font:inherit;font-size:13px;cursor:pointer}
.bb-cp-segbtn:last-child{border-right:0}
.bb-cp-segbtn.on{background:var(--bb-ui-active-bg,#e6f1fb);color:var(--bb-ui-active-fg,#0c447c)}
.bb-cp-presets{display:grid;grid-template-columns:repeat(5,1fr);gap:6px;margin-bottom:12px}
.bb-cp-preset{display:flex;align-items:center;justify-content:center;height:32px;border:0.5px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;background:var(--bb-ui-bg,#fff);cursor:pointer;padding:0}
.bb-cp-preset.on{border-color:var(--bb-ui-active-border,#7fb2ec);background:var(--bb-ui-active-bg,#e6f1fb)}
.bb-cp-preset:disabled{opacity:.35;cursor:default}
.bb-cp-penrow{display:flex;gap:10px;margin-bottom:10px}
.bb-cp-select{flex:1;height:32px;border:0.5px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;background:var(--bb-ui-bg,#fff);color:inherit;font:inherit;font-size:13px;padding:0 8px}
.bb-cp-footer{display:flex;justify-content:flex-end;gap:8px;margin-top:2px}
.bb-cp-btn{height:32px;padding:0 16px;border:0.5px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;background:var(--bb-ui-bg,#fff);color:inherit;font:inherit;font-size:13px;cursor:pointer}
.bb-cp-primary{background:var(--bb-ui-active-bg,#e6f1fb);border-color:var(--bb-ui-active-border,#b5d4f4);color:var(--bb-ui-active-fg,#0c447c)}
`;

const ACTIVE = '#378ADD';
const INACTIVE = '#c9c7be';
const presetLine = (
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  on: number,
) =>
  `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${on ? ACTIVE : INACTIVE}" stroke-width="${on ? 1.6 : 1}" ${on ? '' : 'stroke-dasharray="1.5 1.5"'} stroke-linecap="round"/>`;
const presetIcon = (s: number[]) =>
  `<svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true">` +
  presetLine(3, 3, 17, 3, s[0]) +
  presetLine(17, 3, 17, 17, s[1]) +
  presetLine(3, 17, 17, 17, s[2]) +
  presetLine(3, 3, 3, 17, s[3]) +
  presetLine(3, 10, 17, 10, s[4]) +
  presetLine(10, 3, 10, 17, s[5]) +
  `</svg>`;

const section = (label: string): HTMLDivElement => {
  const sec = document.createElement('div');
  const l = document.createElement('div');
  l.className = 'bb-cp-label';
  l.textContent = label;
  sec.appendChild(l);
  return sec;
};

/**
 * The cell-properties dialog (modal): fill, vertical alignment, and a Word/Docs
 * style borders panel (10 presets + width / style / colour). Borders are
 * action-based — pick the pen, click a preset, and on Apply the host paints
 * that preset across the selected block. Untouched borders are left as-is.
 */
export function openCellProperties({
  initial,
  singleCell,
  onApply,
}: CellPropertiesOptions): void {
  injectStyle('bb-ui-cellprops-styles', STYLE);
  let background = initial.background;
  let vAlign = initial.vAlign;
  let penWidthPt = 1;
  let penStyle: BorderStyle = 'solid';
  let penColor = '#000000';
  let preset: BorderPreset | null = null;

  const root = document.createElement('div');
  root.className = 'bb-cp';

  // ── Fill ──
  const fill = section('Fill');
  const fillBtn = colorButton({
    title: 'Cell fill',
    label: 'Fill',
    clearLabel: 'No fill',
    value: () => background,
    onPick: (c) => {
      background = c;
    },
  });
  fill.append(fillBtn.el);

  // ── Vertical alignment ──
  const valign = section('Vertical alignment');
  const seg = document.createElement('div');
  seg.className = 'bb-cp-seg';
  const vaBtns = new Map<CellPropsResult['vAlign'], HTMLButtonElement>();
  for (const { key, label } of VALIGNS) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'bb-cp-segbtn';
    b.textContent = label;
    b.addEventListener('click', () => {
      vAlign = key;
      syncVAlign();
    });
    seg.appendChild(b);
    vaBtns.set(key, b);
  }
  valign.appendChild(seg);

  // ── Borders ──
  const borders = section('Borders');
  const penRow = document.createElement('div');
  penRow.className = 'bb-cp-penrow';
  const widthSel = document.createElement('select');
  widthSel.className = 'bb-cp-select';
  for (const pt of PT_WIDTHS) {
    const o = document.createElement('option');
    o.value = String(pt);
    o.textContent = `${pt} pt`;
    if (pt === penWidthPt) o.selected = true;
    widthSel.appendChild(o);
  }
  widthSel.addEventListener(
    'change',
    () => (penWidthPt = Number(widthSel.value)),
  );
  const styleSel = document.createElement('select');
  styleSel.className = 'bb-cp-select';
  for (const { key, label } of STYLES) {
    const o = document.createElement('option');
    o.value = key;
    o.textContent = label;
    styleSel.appendChild(o);
  }
  styleSel.addEventListener(
    'change',
    () => (penStyle = styleSel.value as BorderStyle),
  );
  const penBtn = colorButton({
    title: 'Border colour',
    label: 'Colour',
    value: () => penColor,
    onPick: (c) => {
      if (c) penColor = c;
    },
  });
  penRow.append(widthSel, styleSel, penBtn.el);

  const grid = document.createElement('div');
  grid.className = 'bb-cp-presets';
  const presetEls: Array<{ el: HTMLButtonElement; key: BorderPreset }> = [];
  for (const p of PRESETS) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'bb-cp-preset';
    btn.title = p.label;
    btn.setAttribute('aria-label', p.label);
    btn.innerHTML = presetIcon(p.sides);
    btn.disabled = singleCell && p.inside; // Inside* meaningless for 1×1
    btn.addEventListener('click', () => {
      preset = preset === p.key ? null : p.key;
      syncPresets();
    });
    grid.appendChild(btn);
    presetEls.push({ el: btn, key: p.key });
  }
  borders.append(grid, penRow);

  // ── Footer ──
  const footer = document.createElement('div');
  footer.className = 'bb-cp-footer';
  const cancel = document.createElement('button');
  cancel.type = 'button';
  cancel.className = 'bb-cp-btn';
  cancel.textContent = 'Cancel';
  const apply = document.createElement('button');
  apply.type = 'button';
  apply.className = 'bb-cp-btn bb-cp-primary';
  apply.textContent = 'Apply';
  footer.append(cancel, apply);

  root.append(fill, valign, borders, footer);

  function syncVAlign(): void {
    for (const [key, b] of vaBtns) b.classList.toggle('on', vAlign === key);
  }
  function syncPresets(): void {
    for (const { el, key } of presetEls)
      el.classList.toggle('on', preset === key);
  }
  syncVAlign();
  syncPresets();

  const dialog = new Dialog({ title: 'Cell properties', modal: true });
  dialog.setContent(root);
  dialog.onClose(() => dialog.destroy());
  cancel.addEventListener('click', () => dialog.close());
  apply.addEventListener('click', () => {
    const result: CellPropsResult = { background, vAlign };
    if (preset) {
      result.border = {
        preset,
        width: (penWidthPt * 96) / 72,
        style: penStyle,
        color: penColor,
      };
    }
    onApply(result);
    dialog.close();
  });
  dialog.open();
}
