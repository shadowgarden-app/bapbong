import { injectStyle, placeFloating, rovingGrid } from './internal.js';

/**
 * The one colour picker.
 *
 * Before this there were eight separate palettes — two in each shell's
 * toolbar (already drifted apart from each other), two in cell-properties, two
 * in the font dialog — and three unrelated pieces of UI for choosing from
 * them. A colour is the same decision wherever it is made, so it gets one
 * panel and one palette.
 *
 * The panel is `position: fixed` on the body rather than a child of whatever
 * opened it. That is not just tidiness: `.bb-toolbar` is `overflow: hidden`,
 * and the popover it used to render had to live outside the row and compute
 * container-relative coordinates to escape being clipped. A body-level panel
 * has no such problem to solve.
 */

/** Google's palette, which is what a Docs user already has in their hands:
 *  a greyscale row, then seven tints and shades of the same ten hues. */
export const COLOR_PALETTE: readonly (readonly string[])[] = [
  [
    '#000000',
    '#434343',
    '#666666',
    '#999999',
    '#B7B7B7',
    '#CCCCCC',
    '#D9D9D9',
    '#EFEFEF',
    '#F3F3F3',
    '#FFFFFF',
  ],
  [
    '#980000',
    '#FF0000',
    '#FF9900',
    '#FFFF00',
    '#00FF00',
    '#00FFFF',
    '#4A86E8',
    '#0000FF',
    '#9900FF',
    '#FF00FF',
  ],
  [
    '#E6B8AF',
    '#F4CCCC',
    '#FCE5CD',
    '#FFF2CC',
    '#D9EAD3',
    '#D0E0E3',
    '#C9DAF8',
    '#CFE2F3',
    '#D9D2E9',
    '#EAD1DC',
  ],
  [
    '#DD7E6B',
    '#EA9999',
    '#F9CB9C',
    '#FFE599',
    '#B6D7A8',
    '#A2C4C9',
    '#A4C2F4',
    '#9FC5E8',
    '#B4A7D6',
    '#D5A6BD',
  ],
  [
    '#CC4125',
    '#E06666',
    '#F6B26B',
    '#FFD966',
    '#93C47D',
    '#76A5AF',
    '#6D9EEB',
    '#6FA8DC',
    '#8E7CC3',
    '#C27BA0',
  ],
  [
    '#A61C00',
    '#CC0000',
    '#E69138',
    '#F1C232',
    '#6AA84F',
    '#45818E',
    '#3C78D8',
    '#3D85C6',
    '#674EA7',
    '#A64D79',
  ],
  [
    '#85200C',
    '#990000',
    '#B45F06',
    '#BF9000',
    '#38761D',
    '#134F5C',
    '#1155CC',
    '#0B5394',
    '#351C75',
    '#741B47',
  ],
  [
    '#5B0F00',
    '#660000',
    '#783F04',
    '#7F6000',
    '#274E13',
    '#0C343D',
    '#1C4587',
    '#073763',
    '#20124D',
    '#4C1130',
  ],
];

/** Colours the user mixed themselves, newest first. Deliberately module-level
 *  and session-lived: persisting them would mean this package reaching into
 *  storage, which nothing here does, for a convenience nobody asked for. */
const customColors: string[] = [];
const MAX_CUSTOM = 8;

function remember(color: string): void {
  const at = customColors.indexOf(color);
  if (at !== -1) customColors.splice(at, 1);
  customColors.unshift(color);
  if (customColors.length > MAX_CUSTOM) customColors.length = MAX_CUSTOM;
}

const isHex = (v: string) => /^#[0-9a-fA-F]{6}$/.test(v);

const STYLE = `
.bb-cpk{position:fixed;margin:0;max-width:none;max-height:none;z-index:1300;padding:11px;background:var(--bb-ui-menu-bg,#fff);-webkit-backdrop-filter:var(--bb-ui-pop-filter,none);backdrop-filter:var(--bb-ui-pop-filter,none);color:var(--bb-ui-fg,#2c2c2a);border:1px solid var(--bb-ui-pop-border,var(--bb-ui-border,#e3e3e0));border-radius:10px;box-shadow:0 8px 28px rgba(0,0,0,.16);font-family:var(--bb-ui-font,system-ui,-apple-system,sans-serif);box-sizing:border-box}
.bb-cpk *{box-sizing:border-box}
.bb-cpk-grid{display:grid;grid-template-columns:repeat(10,20px);gap:5px}
.bb-cpk-sw{width:20px;height:20px;padding:0;border-radius:50%;border:1px solid rgba(128,128,128,.35);cursor:pointer}
.bb-cpk-sw:focus-visible{outline:none}
.bb-cpk-sw.on,.bb-cpk-sw:focus-visible{box-shadow:0 0 0 2px var(--bb-ui-menu-bg,#fff),0 0 0 4px var(--bb-ui-active-fg,#0c447c)}
.bb-cpk-cap{font-size:11px;letter-spacing:.04em;opacity:.6;margin:13px 0 7px}
.bb-cpk-row{display:flex;gap:6px;align-items:center;flex-wrap:wrap}
.bb-cpk-none{width:100%;display:flex;align-items:center;gap:8px;height:28px;padding:0 7px;margin-bottom:10px;border:0;border-radius:6px;background:transparent;color:inherit;font:inherit;font-size:13px;text-align:left;cursor:pointer}
.bb-cpk-none:hover,.bb-cpk-none:focus-visible{background:var(--bb-ui-hover,#f1efe8);outline:none}
.bb-cpk-none-chip{width:18px;height:18px;border-radius:50%;border:1px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));display:flex;align-items:center;justify-content:center;font-size:12px;opacity:.6;flex:none}
.bb-cpk-plus{width:20px;height:20px;padding:0;border-radius:50%;border:1px dashed var(--bb-ui-border,#d8d6cf);background:transparent;color:inherit;font:inherit;font-size:13px;line-height:1;cursor:pointer;opacity:.75;position:relative;overflow:hidden}
.bb-cpk-plus input{position:absolute;inset:0;opacity:0;cursor:pointer;border:0;padding:0}
.bb-cpk-hex{height:26px;width:96px;border:1px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-radius:6px;background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff));color:inherit;font:inherit;font-size:12px;padding:0 7px;font-family:var(--bb-ui-mono,ui-monospace,monospace)}

.bb-cpk-btn{min-width:30px;height:30px;display:inline-flex;align-items:center;justify-content:center;gap:5px;border:1px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-radius:6px;background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff));color:inherit;font:inherit;font-size:13px;padding:0 7px;cursor:pointer}
.bb-cpk-btn-chip{width:14px;height:14px;border-radius:50%;border:1px solid rgba(128,128,128,.4);flex:none}
.bb-cpk-btn-chip.none{background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff));display:flex;align-items:center;justify-content:center;font-size:10px;opacity:.6}
.bb-cpk-btn-caret{opacity:.5;font-size:10px}
/* Toolbar shape: a glyph with the colour as a bar beneath it. */
.bb-cpk-btn.glyphed{flex-direction:column;gap:1px;min-width:30px;padding:0 6px;border-color:transparent;background:transparent}
.bb-cpk-btn.glyphed:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-cpk-btn-glyph{font-size:13px;line-height:1}
.bb-cpk-btn-bar{width:16px;height:4px;border-radius:1px;border:0.5px solid rgba(128,128,128,.35)}
.bb-cpk::backdrop{background:transparent}
`;

export interface ColorPickerOptions {
  /** The control the panel hangs from. */
  anchor: HTMLElement;
  /** Applied colour — ticked in the grid. */
  value: string | null;
  /** Text for the "no colour" row ("Automatic", "No highlight"). Omit when
   *  the property has no such state, as a border colour does not. */
  clearLabel?: string;
  onPick(color: string | null): void;
}

/** Open the panel. Closes on pick, outside pointer, Esc or scroll. */
export function openColorPicker({
  anchor,
  value,
  clearLabel,
  onPick,
}: ColorPickerOptions): { close(): void } {
  injectStyle('bb-ui-colorpicker-styles', STYLE);
  // A <dialog> rather than a positioned div, because the picker has to open
  // over the Font and Cell dialogs — and a modal <dialog> lives in the TOP
  // LAYER, which no z-index can climb above. Being one itself is the only way
  // to sit on top, and it brings native Esc and focus containment along.
  const panel = document.createElement('dialog');
  panel.className = 'bb-cpk';
  panel.setAttribute('aria-label', 'Colour');

  let closed = false;
  const close = (): void => {
    if (closed) return;
    closed = true;
    window.removeEventListener('scroll', close, true);
    if (panel.open) panel.close();
    panel.remove();
    anchor.focus();
  };
  const pick = (c: string | null): void => {
    if (c && !COLOR_PALETTE.some((row) => row.includes(c))) remember(c);
    onPick(c);
    close();
  };

  if (clearLabel !== undefined) {
    const none = document.createElement('button');
    none.type = 'button';
    none.className = 'bb-cpk-none';
    const chip = document.createElement('span');
    chip.className = 'bb-cpk-none-chip';
    chip.textContent = '⦸';
    none.append(chip, document.createTextNode(clearLabel));
    none.addEventListener('click', () => pick(null));
    panel.append(none);
  }

  // ── the grid ────────────────────────────────────────────────────
  // A real grid, not a pile of buttons: eighty swatches are unusable from the
  // keyboard without one, and the popover this replaces had no key handling
  // at all. Roving tabindex + arrows, so Tab passes over the whole grid.
  const grid = document.createElement('div');
  grid.className = 'bb-cpk-grid';
  grid.setAttribute('role', 'grid');
  const cells: HTMLButtonElement[][] = [];
  COLOR_PALETTE.forEach((row) => {
    const rowCells: HTMLButtonElement[] = [];
    row.forEach((color) => {
      const sw = document.createElement('button');
      sw.type = 'button';
      sw.className = 'bb-cpk-sw';
      sw.style.background = color;
      sw.title = color;
      sw.setAttribute('aria-label', color);
      if (color.toUpperCase() === value?.toUpperCase()) sw.classList.add('on');
      sw.addEventListener('click', () => pick(color));
      grid.append(sw);
      rowCells.push(sw);
    });
    cells.push(rowCells);
  });
  panel.append(grid);
  const nav = rovingGrid(cells); // arrows + roving tabindex, shared with the symbol grid

  // ── custom ──────────────────────────────────────────────────────
  panel.append(
    Object.assign(document.createElement('div'), {
      className: 'bb-cpk-cap',
      textContent: 'CUSTOM',
    }),
  );
  const customRow = document.createElement('div');
  customRow.className = 'bb-cpk-row';
  for (const c of customColors) {
    const sw = document.createElement('button');
    sw.type = 'button';
    sw.className = 'bb-cpk-sw';
    sw.style.background = c;
    sw.title = c;
    sw.setAttribute('aria-label', c);
    if (c.toUpperCase() === value?.toUpperCase()) sw.classList.add('on');
    sw.addEventListener('click', () => pick(c));
    customRow.append(sw);
  }
  // `+` is the platform picker, not a hex field pretending to be one.
  const plus = document.createElement('button');
  plus.type = 'button';
  plus.className = 'bb-cpk-plus';
  plus.title = 'Custom colour';
  plus.append(document.createTextNode('+'));
  const native = document.createElement('input');
  native.type = 'color';
  native.value = value && isHex(value) ? value : '#000000';
  native.addEventListener('input', () => pick(native.value.toUpperCase()));
  plus.append(native);
  customRow.append(plus);

  // The hex field stays: someone holding a brand colour pastes it faster than
  // they can find it in a picker.
  const hex = document.createElement('input');
  hex.type = 'text';
  hex.className = 'bb-cpk-hex';
  hex.placeholder = '#RRGGBB';
  hex.spellcheck = false;
  hex.value = value ?? '';
  hex.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    const v = hex.value.trim();
    if (isHex(v)) pick(v.toUpperCase());
  });
  customRow.append(hex);
  panel.append(customRow);

  // ── mount + dismiss ─────────────────────────────────────────────
  document.body.append(panel);
  panel.showModal();
  const r = anchor.getBoundingClientRect();
  placeFloating(panel, { x: r.left, y: r.top, height: r.height });
  // Esc and the ✕-less backdrop are native; `close` fires for both, so the
  // cleanup hangs off that rather than off a keydown listener of our own.
  panel.addEventListener('close', close);
  // A click on the backdrop lands on the dialog element itself. Compare
  // against the box so a click in the panel's own padding does not close it.
  panel.addEventListener('pointerdown', (e) => {
    const b = panel.getBoundingClientRect();
    const outside =
      e.clientX < b.left ||
      e.clientX > b.right ||
      e.clientY < b.top ||
      e.clientY > b.bottom;
    if (outside) close();
  });
  window.addEventListener('scroll', close, true);

  // Open on the applied colour when there is one, so arrows start from where
  // the user already is rather than the top-left corner.
  const onRow = cells.findIndex((row) =>
    row.some((el) => el.classList.contains('on')),
  );
  const onCol =
    onRow >= 0
      ? cells[onRow].findIndex((el) => el.classList.contains('on'))
      : 0;
  nav.focusCell(Math.max(0, onRow), Math.max(0, onCol));

  return { close };
}

export interface ColorButtonOptions {
  /** Current colour, read fresh on every sync — including once during
   *  construction, so it must be safe to call before anything is loaded. */
  value(): string | null;
  onPick(color: string | null): void;
  title: string;
  clearLabel?: string;
  /** Toolbar shape: this glyph above a bar of the colour. Omit for the
   *  dialog shape, a chip beside a caret. */
  glyph?: string;
  /** Extra label inside the button (dialogs use it; the toolbar does not). */
  label?: string;
}

/** A button showing the current colour that opens {@link openColorPicker}. */
export function colorButton(o: ColorButtonOptions): {
  el: HTMLButtonElement;
  sync(): void;
} {
  injectStyle('bb-ui-colorpicker-styles', STYLE);
  const el = document.createElement('button');
  el.type = 'button';
  el.className = `bb-cpk-btn${o.glyph ? ' glyphed' : ''}`;
  el.title = o.title;
  el.setAttribute('aria-label', o.title);
  el.setAttribute('aria-haspopup', 'dialog');

  let chip: HTMLElement;
  if (o.glyph) {
    const g = document.createElement('span');
    g.className = 'bb-cpk-btn-glyph';
    g.innerHTML = o.glyph; // text or inline SVG, host-trusted (as before)
    chip = document.createElement('span');
    chip.className = 'bb-cpk-btn-bar';
    el.append(g, chip);
  } else {
    chip = document.createElement('span');
    chip.className = 'bb-cpk-btn-chip';
    el.append(chip);
    if (o.label) el.append(document.createTextNode(o.label));
    const caret = document.createElement('span');
    caret.className = 'bb-cpk-btn-caret';
    caret.textContent = '▾';
    el.append(caret);
  }

  function sync(): void {
    const v = o.value();
    chip.style.background = v ?? 'transparent';
    chip.classList.toggle('none', !v && !o.glyph);
    chip.textContent = !v && !o.glyph ? '⦸' : '';
  }

  // Keep the editor's selection: focusing the button would collapse it.
  el.addEventListener('mousedown', (e) => e.preventDefault());
  el.addEventListener('click', () => {
    openColorPicker({
      anchor: el,
      value: o.value(),
      clearLabel: o.clearLabel,
      onPick: (c) => {
        o.onPick(c);
        sync();
      },
    });
  });
  sync();
  return { el, sync };
}
