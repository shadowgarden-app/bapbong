import { Dialog } from './dialog.js';
import { injectStyle } from './internal.js';

/**
 * Layout ▸ Page size / Margins: Word-style preset flyouts (a preview icon, the
 * preset name, and its measurements) plus the two "custom…" dialogs.
 *
 * Presentation only — the preset tables live in bapbong-commands and are passed
 * in as data, so this package keeps its single contracts dependency and the
 * widgets stay reusable by any shell.
 *
 * Measurements are shown in centimetres (what Word's Page Setup shows) while
 * the model is CSS px @96dpi; the conversion lives here, at the only boundary
 * that speaks cm.
 */

const PX_PER_CM = 96 / 2.54;

/** px → cm. */
export const pxToCm = (px: number): number => px / PX_PER_CM;
/** cm → px (rounded — the model is integer px). */
export const cmToPx = (cm: number): number => Math.round(cm * PX_PER_CM);

/** "2.54 cm" — up to 3 decimals, trailing zeros trimmed. */
function fmtCm(cm: number): string {
  return `${Number(cm.toFixed(3))} cm`;
}

export interface PaperChoice {
  /** Stable key echoed back to `onPick` (e.g. a PaperSize name). */
  key: string;
  label: string;
  /** Nominal portrait size in cm — the caption. Exact values, not derived
   *  from px (rounding would print A4 as 21.006 cm). */
  cm: readonly [number, number];
  /** Portrait size in px — drives the preview icon's aspect ratio. */
  px: readonly [number, number];
  active?: boolean;
}

export interface MarginChoice {
  key: string;
  label: string;
  /** Margins in px; shown in cm and drawn as the preview's content box. */
  margin: PageMargin;
  active?: boolean;
}

export interface PageMargin {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

const STYLE = `
.bb-ps{display:flex;flex-direction:column;min-width:250px;max-height:min(70vh,460px);overflow-y:auto}
.bb-ps-row{display:flex;align-items:center;gap:11px;width:100%;padding:7px 10px;border:0;border-radius:6px;background:transparent;color:inherit;font:inherit;text-align:left;cursor:pointer}
.bb-ps-row:hover,.bb-ps-row:focus{background:var(--bb-ui-hover,#f1efe8);outline:none}
.bb-ps-row[aria-checked="true"]{background:var(--bb-ui-active-bg,#e6f1fb)}
.bb-ps-icon{flex:none;width:34px;height:42px;display:block}
.bb-ps-text{flex:1 1 auto;min-width:0}
.bb-ps-name{font-size:13px;font-weight:600;line-height:1.3}
.bb-ps-dims{font-size:11.5px;opacity:.68;line-height:1.45;white-space:nowrap}
.bb-ps-grid{display:grid;grid-template-columns:auto auto;gap:0 14px;justify-content:start}
.bb-ps-sep{height:1px;margin:4px 6px;background:var(--bb-ui-border,#e3e3e0)}
.bb-pd{display:flex;flex-direction:column;gap:14px;min-width:300px}
.bb-pd-fields{display:grid;gap:10px 18px}
.bb-pd-field{display:flex;align-items:center;gap:8px}
.bb-pd-label{flex:0 0 74px;font-size:13px}
.bb-pd-input{flex:1 1 auto;min-width:0;width:100%;height:30px;padding:0 8px;border:1px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;background:var(--bb-ui-bg,#fff);color:inherit;font:inherit;font-size:13px}
.bb-pd-input:focus{outline:2px solid var(--bb-ui-active-border,#7fb2ec);outline-offset:-1px}
.bb-pd-actions{display:flex;justify-content:flex-end;gap:8px}
`;

const SVG_NS = 'http://www.w3.org/2000/svg';

function svgEl<K extends keyof SVGElementTagNameMap>(
  name: K,
  attrs: Record<string, string | number>,
): SVGElementTagNameMap[K] {
  const el = document.createElementNS(SVG_NS, name);
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
  return el;
}

/** The 34×42 preview frame with a page outline sized to `ratio` (w/h). */
function pagePreview(ratio: number): SVGSVGElement {
  const svg = svgEl('svg', { viewBox: '0 0 34 42', class: 'bb-ps-icon' });
  // Fit the page inside a 26×36 box, centred — so Legal reads taller and
  // narrower than Letter at a glance.
  const maxW = 26;
  const maxH = 36;
  const w = Math.min(maxW, maxH * ratio);
  const h = w / ratio;
  svg.appendChild(
    svgEl('rect', {
      x: (34 - w) / 2,
      y: (42 - h) / 2,
      width: w,
      height: h,
      rx: 1,
      fill: 'var(--bb-ui-bg,#fff)',
      stroke: 'var(--bb-ui-fg,#2c2c2a)',
      'stroke-opacity': 0.45,
      'stroke-width': 1,
    }),
  );
  return svg;
}

/** A page outline with its content box drawn inside, both to scale. */
function marginPreview(
  margin: PageMargin,
  page: { width: number; height: number },
): SVGSVGElement {
  const svg = pagePreview(page.width / page.height);
  const outline = svg.firstChild as SVGRectElement;
  const px = Number(outline.getAttribute('x'));
  const py = Number(outline.getAttribute('y'));
  const pw = Number(outline.getAttribute('width'));
  const ph = Number(outline.getAttribute('height'));
  const sx = pw / page.width;
  const sy = ph / page.height;
  const cw = Math.max(2, pw - (margin.left + margin.right) * sx);
  const ch = Math.max(2, ph - (margin.top + margin.bottom) * sy);
  svg.appendChild(
    svgEl('rect', {
      x: px + margin.left * sx,
      y: py + margin.top * sy,
      width: cw,
      height: ch,
      fill: 'var(--bb-ui-active-border,#7fb2ec)',
      'fill-opacity': 0.28,
      stroke: 'var(--bb-ui-active-fg,#0c447c)',
      'stroke-opacity': 0.55,
      'stroke-width': 0.8,
    }),
  );
  return svg;
}

/** One flyout row: preview icon + name + caption element(s). */
function presetRow(
  icon: SVGSVGElement,
  name: string,
  caption: HTMLElement,
  active: boolean,
  onPick: () => void,
): HTMLButtonElement {
  const row = document.createElement('button');
  row.type = 'button';
  row.className = 'bb-ps-row';
  row.setAttribute('role', 'menuitemradio');
  row.setAttribute('aria-checked', String(active));
  const text = document.createElement('span');
  text.className = 'bb-ps-text';
  const nameEl = document.createElement('span');
  nameEl.className = 'bb-ps-name';
  nameEl.textContent = name;
  text.append(nameEl, caption);
  row.append(icon, text);
  // Keep focus in the editor while the menu is open (matches the grid picker).
  row.addEventListener('mousedown', (e) => e.preventDefault());
  row.addEventListener('click', onPick);
  return row;
}

function captionLine(text: string): HTMLElement {
  const el = document.createElement('span');
  el.className = 'bb-ps-dims';
  el.style.display = 'block';
  el.textContent = text;
  return el;
}

/** The trailing "Custom …" row (title + description, no measurements). */
function customRow(
  icon: SVGSVGElement,
  title: string,
  description: string,
  onPick: () => void,
): HTMLButtonElement {
  // A small badge marks it as the "define your own" entry, like Word's.
  icon.appendChild(
    svgEl('circle', {
      cx: 26,
      cy: 34,
      r: 5.5,
      fill: 'var(--bb-ui-active-fg,#0c447c)',
    }),
  );
  const plus = svgEl('path', {
    d: 'M26 31.2v5.6M23.2 34h5.6',
    stroke: 'var(--bb-ui-bg,#fff)',
    'stroke-width': 1.4,
    'stroke-linecap': 'round',
  });
  icon.appendChild(plus);
  return presetRow(icon, title, captionLine(description), false, onPick);
}

export interface PageSizePickerOptions {
  items: readonly PaperChoice[];
  /** Called with the chosen item's `key`. */
  onPick: (key: string) => void;
  /** Called for the trailing "Custom page size" row. */
  onCustom: () => void;
}

/**
 * Layout ▸ Page size flyout: one row per paper preset (preview, name, size in
 * cm) plus a custom row. Returns a detached element for a menu widget slot.
 */
export function pageSizePicker(options: PageSizePickerOptions): HTMLElement {
  injectStyle('bb-ui-pagesetup-styles', STYLE);
  const root = document.createElement('div');
  root.className = 'bb-ps';
  root.setAttribute('role', 'menu');

  for (const item of options.items) {
    root.appendChild(
      presetRow(
        pagePreview(item.px[0] / item.px[1]),
        item.label,
        captionLine(`${fmtCm(item.cm[0])} x ${fmtCm(item.cm[1])}`),
        !!item.active,
        () => options.onPick(item.key),
      ),
    );
  }

  const sep = document.createElement('div');
  sep.className = 'bb-ps-sep';
  root.appendChild(sep);
  root.appendChild(
    customRow(
      pagePreview(0.773),
      'Custom page size',
      'Define custom page size',
      options.onCustom,
    ),
  );
  return root;
}

export interface OrientationPickerOptions {
  /** Current page size in px — the previews are drawn at its aspect ratio. */
  page: { width: number; height: number };
  active: 'portrait' | 'landscape';
  onPick: (orientation: 'portrait' | 'landscape') => void;
}

/**
 * Layout ▸ Orientation flyout. Same row shape as the other two pickers, so all
 * three entries of the Layout menu read as one family.
 */
export function orientationPicker(
  options: OrientationPickerOptions,
): HTMLElement {
  injectStyle('bb-ui-pagesetup-styles', STYLE);
  const root = document.createElement('div');
  root.className = 'bb-ps';
  root.setAttribute('role', 'menu');
  // Normalize to portrait so both previews are drawn from the same paper.
  const short = Math.min(options.page.width, options.page.height);
  const long = Math.max(options.page.width, options.page.height);
  for (const [key, label, ratio] of [
    ['portrait', 'Portrait', short / long],
    ['landscape', 'Landscape', long / short],
  ] as const) {
    root.appendChild(
      presetRow(
        pagePreview(ratio),
        label,
        captionLine(
          `${fmtCm(pxToCm(key === 'portrait' ? short : long))} x ${fmtCm(
            pxToCm(key === 'portrait' ? long : short),
          )}`,
        ),
        options.active === key,
        () => options.onPick(key),
      ),
    );
  }
  return root;
}

export interface MarginPickerOptions {
  items: readonly MarginChoice[];
  /** Page size in px — the previews draw each preset's content box to scale. */
  page: { width: number; height: number };
  onPick: (key: string) => void;
  onCustom: () => void;
}

/**
 * Layout ▸ Margins flyout: one row per preset (a preview showing the content
 * box to scale, the name, and the four edges in cm) plus a custom row.
 */
export function marginPresetPicker(options: MarginPickerOptions): HTMLElement {
  injectStyle('bb-ui-pagesetup-styles', STYLE);
  const root = document.createElement('div');
  root.className = 'bb-ps';
  root.setAttribute('role', 'menu');

  for (const item of options.items) {
    const grid = document.createElement('span');
    grid.className = 'bb-ps-dims bb-ps-grid';
    for (const [label, px] of [
      ['Top:', item.margin.top],
      ['Bottom:', item.margin.bottom],
      ['Left:', item.margin.left],
      ['Right:', item.margin.right],
    ] as const) {
      const cell = document.createElement('span');
      cell.textContent = `${label} ${fmtCm(pxToCm(px))}`;
      grid.appendChild(cell);
    }
    root.appendChild(
      presetRow(
        marginPreview(item.margin, options.page),
        item.label,
        grid,
        !!item.active,
        () => options.onPick(item.key),
      ),
    );
  }

  const sep = document.createElement('div');
  sep.className = 'bb-ps-sep';
  root.appendChild(sep);
  root.appendChild(
    customRow(
      marginPreview(
        { top: 96, right: 96, bottom: 96, left: 96 },
        options.page,
      ),
      'Custom margins',
      'Define custom margins',
      options.onCustom,
    ),
  );
  return root;
}

/** A labelled cm input. `min` guards the layout engine's floor. */
function cmField(
  label: string,
  valuePx: number,
  min: number,
): { row: HTMLElement; input: HTMLInputElement } {
  const row = document.createElement('label');
  row.className = 'bb-pd-field';
  const name = document.createElement('span');
  name.className = 'bb-pd-label';
  name.textContent = label;
  const input = document.createElement('input');
  input.className = 'bb-pd-input';
  input.type = 'number';
  input.step = '0.01';
  input.min = String(Number(pxToCm(min).toFixed(3)));
  input.value = String(Number(pxToCm(valuePx).toFixed(3)));
  row.append(name, input);
  return { row, input };
}

/** The OK / Cancel strip shared by both dialogs. */
function dialogActions(): {
  el: HTMLElement;
  ok: HTMLButtonElement;
  cancel: HTMLButtonElement;
} {
  const el = document.createElement('div');
  el.className = 'bb-pd-actions';
  const cancel = document.createElement('button');
  cancel.type = 'button';
  cancel.className = 'bb-prompt-btn';
  cancel.textContent = 'Cancel';
  const ok = document.createElement('button');
  ok.type = 'submit';
  ok.className = 'bb-prompt-btn';
  ok.dataset['primary'] = 'true';
  ok.textContent = 'OK';
  el.append(cancel, ok);
  return { el, ok, cancel };
}

/** Read an input as px, falling back to `fallbackPx` on a blank/NaN entry. */
function readPx(input: HTMLInputElement, fallbackPx: number): number {
  const n = Number(input.value);
  return Number.isFinite(n) && n > 0 ? cmToPx(n) : fallbackPx;
}

export interface PageSizeDialogOptions {
  /** Current size in px. */
  initial: { width: number; height: number };
  /** Applied size in px. */
  onApply: (size: { width: number; height: number }) => void;
}

/** Layout ▸ Page size ▸ Custom page size — width/height in cm. */
export function openPageSizeDialog(options: PageSizeDialogOptions): void {
  injectStyle('bb-ui-pagesetup-styles', STYLE);
  const dialog = new Dialog({ title: 'Paper size', modal: true });
  const form = document.createElement('form');
  form.className = 'bb-pd';
  const fields = document.createElement('div');
  fields.className = 'bb-pd-fields';
  const width = cmField('Width:', options.initial.width, 96);
  const height = cmField('Height:', options.initial.height, 96);
  fields.append(width.row, height.row);
  const actions = dialogActions();
  form.append(fields, actions.el);

  dialog.setContent(form);
  dialog.onClose(() => dialog.destroy());
  actions.cancel.addEventListener('click', () => dialog.close());
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    options.onApply({
      width: readPx(width.input, options.initial.width),
      height: readPx(height.input, options.initial.height),
    });
    dialog.close();
  });
  dialog.open();
  width.input.focus();
  width.input.select();
}

export interface MarginsDialogOptions {
  /** Current margins in px. */
  initial: PageMargin;
  /** Applied margins in px. */
  onApply: (margin: PageMargin) => void;
}

/** Layout ▸ Margins ▸ Custom margins — the four edges in cm. */
export function openMarginsDialog(options: MarginsDialogOptions): void {
  injectStyle('bb-ui-pagesetup-styles', STYLE);
  const dialog = new Dialog({ title: 'Margins', modal: true });
  const form = document.createElement('form');
  form.className = 'bb-pd';
  const fields = document.createElement('div');
  fields.className = 'bb-pd-fields';
  // Two columns, mirroring Word: Top/Left down the first, Bottom/Right the
  // second — so opposite edges sit side by side.
  fields.style.gridTemplateColumns = '1fr 1fr';
  const top = cmField('Top:', options.initial.top, 0);
  const left = cmField('Left:', options.initial.left, 0);
  const bottom = cmField('Bottom:', options.initial.bottom, 0);
  const right = cmField('Right:', options.initial.right, 0);
  fields.append(top.row, left.row, bottom.row, right.row);
  const actions = dialogActions();
  form.append(fields, actions.el);

  dialog.setContent(form);
  dialog.onClose(() => dialog.destroy());
  actions.cancel.addEventListener('click', () => dialog.close());
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    // 0 is a legitimate margin, so these read straight (no positive guard).
    const px = (input: HTMLInputElement, fallback: number) => {
      const n = Number(input.value);
      return Number.isFinite(n) && n >= 0 ? cmToPx(n) : fallback;
    };
    options.onApply({
      top: px(top.input, options.initial.top),
      right: px(right.input, options.initial.right),
      bottom: px(bottom.input, options.initial.bottom),
      left: px(left.input, options.initial.left),
    });
    dialog.close();
  });
  dialog.open();
  top.input.focus();
  top.input.select();
}
