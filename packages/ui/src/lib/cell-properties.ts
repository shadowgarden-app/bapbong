import { Dialog } from './dialog.js';
import { injectStyle } from './internal.js';

export interface CellBorders {
  top: boolean;
  right: boolean;
  bottom: boolean;
  left: boolean;
}

/** Cell properties the dialog edits (host maps these onto the cell schema:
 *  `vAlign: 'top'` → null, borders → the per-side override). */
export interface CellProps {
  background: string | null;
  vAlign: 'top' | 'center' | 'bottom';
  borders: CellBorders;
}

export interface CellPropertiesOptions {
  /** Pre-fill from the current cell. */
  initial: CellProps;
  /** Called with the chosen props when Apply is pressed. */
  onApply: (props: CellProps) => void;
}

const PRESETS = ['#FCEBEB', '#FAEEDA', '#EAF3DE', '#E6F1FB', '#EEEDFE', '#FBEAF0', '#E1F5EE', '#F1EFE8', '#D3D1C7'];
const SIDES = ['top', 'right', 'bottom', 'left'] as const;
const VALIGNS: ReadonlyArray<{ key: CellProps['vAlign']; label: string }> = [
  { key: 'top', label: 'Top' },
  { key: 'center', label: 'Middle' },
  { key: 'bottom', label: 'Bottom' },
];

const STYLE = `
.bb-cp{display:flex;flex-direction:column;gap:16px;min-width:300px;color:var(--bb-ui-fg,#2c2c2a)}
.bb-cp *{box-sizing:border-box}
.bb-cp-label{font-size:12px;opacity:.7;margin-bottom:8px}
.bb-cp-swatches{display:flex;flex-wrap:wrap;gap:7px}
.bb-cp-swatch{width:26px;height:26px;border-radius:6px;border:0.5px solid var(--bb-ui-border,#d8d6cf);padding:0;cursor:pointer}
.bb-cp-swatch.on{box-shadow:0 0 0 2px var(--bb-ui-active-fg,#0c447c)}
.bb-cp-none{display:flex;align-items:center;justify-content:center;background:var(--bb-ui-bg,#fff);font-size:15px;color:#b4b2a9}
.bb-cp-hexrow{display:flex;align-items:center;gap:8px;margin-top:10px}
.bb-cp-hexpreview{width:26px;height:26px;border-radius:6px;border:0.5px solid var(--bb-ui-border,#d8d6cf);flex:none}
.bb-cp-hex{height:30px;width:120px;font-size:13px;padding:0 8px;border:0.5px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;background:var(--bb-ui-bg,#fff);color:inherit}
.bb-cp-seg{display:inline-flex;border:0.5px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;overflow:hidden}
.bb-cp-segbtn{height:32px;padding:0 16px;border:0;border-right:0.5px solid var(--bb-ui-border,#e3e3e0);background:transparent;color:inherit;font:inherit;font-size:13px;cursor:pointer}
.bb-cp-segbtn:last-child{border-right:0}
.bb-cp-segbtn.on{background:var(--bb-ui-active-bg,#e6f1fb);color:var(--bb-ui-active-fg,#0c447c)}
.bb-cp-bdhead{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px}
.bb-cp-mini{height:26px;padding:0 10px;font-size:12px;border:0.5px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;background:var(--bb-ui-bg,#fff);color:inherit;cursor:pointer}
.bb-cp-box{position:relative;width:80px;height:60px}
.bb-cp-box-inner{position:absolute;inset:0;border:0.5px dashed var(--bb-ui-border,#d8d6cf);border-radius:4px;pointer-events:none}
.bb-cp-edge{position:absolute;border:0;padding:0;border-radius:3px;background:var(--bb-ui-border,#c9c7be);cursor:pointer}
.bb-cp-edge.on{background:var(--bb-ui-active-fg,#0c447c)}
.bb-cp-top{top:-3px;left:12px;right:12px;height:6px}
.bb-cp-bottom{bottom:-3px;left:12px;right:12px;height:6px}
.bb-cp-left{left:-3px;top:12px;bottom:12px;width:6px}
.bb-cp-right{right:-3px;top:12px;bottom:12px;width:6px}
.bb-cp-footer{display:flex;justify-content:flex-end;gap:8px;margin-top:2px}
.bb-cp-btn{height:32px;padding:0 16px;border:0.5px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;background:var(--bb-ui-bg,#fff);color:inherit;font:inherit;font-size:13px;cursor:pointer}
.bb-cp-primary{background:var(--bb-ui-active-bg,#e6f1fb);border-color:var(--bb-ui-active-border,#b5d4f4);color:var(--bb-ui-active-fg,#0c447c)}
`;

const section = (label: string): HTMLDivElement => {
  const sec = document.createElement('div');
  const l = document.createElement('div');
  l.className = 'bb-cp-label';
  l.textContent = label;
  sec.appendChild(l);
  return sec;
};

/**
 * Open the cell-properties dialog (modal, built on {@link Dialog}): fill colour,
 * vertical alignment, and per-side borders. Pre-filled from `initial`; calls
 * `onApply` with the chosen props on Apply. The host maps the result onto the
 * cell schema and applies it across the selected cells.
 */
export function openCellProperties({ initial, onApply }: CellPropertiesOptions): void {
  injectStyle('bb-ui-cellprops-styles', STYLE);
  const state: CellProps = { background: initial.background, vAlign: initial.vAlign, borders: { ...initial.borders } };

  const root = document.createElement('div');
  root.className = 'bb-cp';

  // ── Fill ──
  const fill = section('Fill');
  const swatches = document.createElement('div');
  swatches.className = 'bb-cp-swatches';
  const swatchEls: Array<{ el: HTMLElement; color: string | null }> = [];
  const none = document.createElement('button');
  none.type = 'button';
  none.className = 'bb-cp-swatch bb-cp-none';
  none.title = 'No fill';
  none.textContent = '⦸';
  none.addEventListener('click', () => {
    state.background = null;
    sync();
  });
  swatches.appendChild(none);
  swatchEls.push({ el: none, color: null });
  for (const color of PRESETS) {
    const sw = document.createElement('button');
    sw.type = 'button';
    sw.className = 'bb-cp-swatch';
    sw.style.background = color;
    sw.title = color;
    sw.addEventListener('click', () => {
      state.background = color;
      sync();
    });
    swatches.appendChild(sw);
    swatchEls.push({ el: sw, color });
  }
  const hexRow = document.createElement('div');
  hexRow.className = 'bb-cp-hexrow';
  const hexPreview = document.createElement('span');
  hexPreview.className = 'bb-cp-hexpreview';
  const hex = document.createElement('input');
  hex.type = 'text';
  hex.className = 'bb-cp-hex';
  hex.placeholder = '#RRGGBB';
  hex.addEventListener('input', () => {
    const v = hex.value.trim();
    if (/^#[0-9a-fA-F]{6}$/.test(v)) {
      state.background = v;
      sync();
    }
  });
  hexRow.append(hexPreview, hex);
  fill.append(swatches, hexRow);

  // ── Vertical alignment ──
  const valign = section('Vertical alignment');
  const seg = document.createElement('div');
  seg.className = 'bb-cp-seg';
  const vaBtns = new Map<CellProps['vAlign'], HTMLButtonElement>();
  for (const { key, label } of VALIGNS) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'bb-cp-segbtn';
    b.textContent = label;
    b.addEventListener('click', () => {
      state.vAlign = key;
      sync();
    });
    seg.appendChild(b);
    vaBtns.set(key, b);
  }
  valign.appendChild(seg);

  // ── Borders ──
  const borders = document.createElement('div');
  const head = document.createElement('div');
  head.className = 'bb-cp-bdhead';
  const bl = document.createElement('span');
  bl.className = 'bb-cp-label';
  bl.style.marginBottom = '0';
  bl.textContent = 'Borders';
  const quick = document.createElement('span');
  const setAll = (on: boolean) => () => {
    state.borders = { top: on, right: on, bottom: on, left: on };
    sync();
  };
  const allBtn = document.createElement('button');
  allBtn.type = 'button';
  allBtn.className = 'bb-cp-mini';
  allBtn.textContent = 'All';
  allBtn.style.marginRight = '6px';
  allBtn.addEventListener('click', setAll(true));
  const noneBtn = document.createElement('button');
  noneBtn.type = 'button';
  noneBtn.className = 'bb-cp-mini';
  noneBtn.textContent = 'None';
  noneBtn.addEventListener('click', setAll(false));
  quick.append(allBtn, noneBtn);
  head.append(bl, quick);
  const box = document.createElement('div');
  box.className = 'bb-cp-box';
  const inner = document.createElement('div');
  inner.className = 'bb-cp-box-inner';
  box.appendChild(inner);
  const edgeEls = new Map<(typeof SIDES)[number], HTMLButtonElement>();
  for (const side of SIDES) {
    const e = document.createElement('button');
    e.type = 'button';
    e.className = `bb-cp-edge bb-cp-${side}`;
    e.title = `${side} border`;
    e.addEventListener('click', () => {
      state.borders[side] = !state.borders[side];
      sync();
    });
    box.appendChild(e);
    edgeEls.set(side, e);
  }
  borders.append(head, box);

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

  function sync(): void {
    for (const { el, color } of swatchEls) el.classList.toggle('on', color === state.background);
    hexPreview.style.background = state.background ?? 'transparent';
    if (document.activeElement !== hex) hex.value = state.background ?? '';
    for (const [key, b] of vaBtns) b.classList.toggle('on', state.vAlign === key);
    for (const [side, e] of edgeEls) e.classList.toggle('on', state.borders[side]);
  }
  sync();

  const dialog = new Dialog({ title: 'Cell properties', modal: true });
  dialog.setContent(root);
  dialog.onClose(() => dialog.destroy());
  cancel.addEventListener('click', () => dialog.close());
  apply.addEventListener('click', () => {
    onApply({ background: state.background, vAlign: state.vAlign, borders: { ...state.borders } });
    dialog.close();
  });
  dialog.open();
}
