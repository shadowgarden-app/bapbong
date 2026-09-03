import {
  cellStyleLayer,
  type BorderSide,
  type ResolvedTableStyle,
  type TableBorders,
  type TableLook,
} from '@shadow-garden/bapbong-contracts';
import { Dialog } from './dialog.js';
import { injectStyle } from './internal.js';

/**
 * The Table style panel — Word's Table Design surface as one non-modal
 * floating dialog: the six tblLook gates ("Table style options") and a
 * gallery of styles whose thumbnails are drawn from the SAME resolved sheet
 * and the SAME cellStyleLayer the renderer uses, so a preview can never
 * disagree with the page.
 *
 * The host owns the data and the commands (applyTableStyle / setTableLook);
 * the panel renders and reports. All three entry points — Format ▸ Table ▸
 * Table style…, the contextual toolbar button, the table context menu — open
 * this one instance.
 */

export interface TableStyleGalleryItem {
  id: string;
  /** Human name for the tile tooltip + footer ("Light Grid · Accent 3"). */
  name: string;
  style: ResolvedTableStyle;
}

export interface TableStylePanelOptions {
  /** Gallery content, re-read on every {@link TableStylePanelHandle.refresh}:
   *  typically the document's own sheet plus the host's built-in catalog. */
  styles: () => TableStyleGalleryItem[];
  /** The enclosing table's pair, or null when the caret left every table. */
  current: () => { styleId: string | null; look: TableLook } | null;
  /** A gallery pick (null = the "None" tile). */
  onPick: (item: TableStyleGalleryItem | null) => void;
  /** A gate checkbox flip. */
  onLook: (patch: Partial<TableLook>) => void;
  /** Where to float (see Dialog.anchor / panelAnchor). */
  anchor?: () => DOMRect | null;
}

export interface TableStylePanelHandle {
  open(): void;
  close(): void;
  /** Re-read styles()/current() and redraw. Cheap no-op when nothing moved. */
  refresh(): void;
  destroy(): void;
  readonly isOpen: boolean;
}

const GATES: Array<{ key: keyof TableLook; label: string }> = [
  { key: 'firstRow', label: 'Header row' },
  { key: 'lastRow', label: 'Total row' },
  { key: 'hBand', label: 'Banded rows' },
  { key: 'firstCol', label: 'First column' },
  { key: 'lastCol', label: 'Last column' },
  { key: 'vBand', label: 'Banded columns' },
];

const THUMB_ROWS = 4;
const THUMB_COLS = 5;

const STYLE = `
.bb-tsp{display:flex;gap:16px;width:432px;max-width:100%;color:var(--bb-ui-fg,#2c2c2a)}
.bb-tsp *{box-sizing:border-box}
.bb-tsp-opts{flex:0 0 128px;display:flex;flex-direction:column;gap:2px}
.bb-tsp-lbl{font-size:12px;opacity:.7;margin-bottom:4px}
.bb-tsp-chk{display:flex;align-items:center;gap:7px;font-size:13px;padding:3px 0;cursor:pointer}
.bb-tsp-chk input{margin:0}
.bb-tsp-chk.is-disabled{opacity:.4;cursor:default}
.bb-tsp-sep{height:1px;margin:6px 0;background:var(--bb-ui-border,#e3e3e0)}
.bb-tsp-main{flex:1;display:flex;flex-direction:column;gap:8px;min-width:0}
.bb-tsp-gallery{display:flex;flex-wrap:wrap;gap:6px}
.bb-tsp-tile{padding:3px;border:1px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-radius:6px;background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff));cursor:pointer;line-height:0}
.bb-tsp-tile:hover{border-color:var(--bb-ui-fg,#2c2c2a)}
.bb-tsp-tile.is-selected{padding:2px;border:2px solid var(--bb-ui-active-fg,#0c447c)}
.bb-tsp-thumb{display:grid;grid-template-columns:repeat(${THUMB_COLS},12px);grid-auto-rows:8px;gap:0}
.bb-tsp-none{display:flex;align-items:center;justify-content:center;width:60px;height:32px;font-size:11px;opacity:.7;line-height:1.1}
.bb-tsp-empty{font-size:13px;opacity:.65;padding:8px 0}
`;

/** One border side's CSS, `false`/absent → none. */
function edgeCss(side: BorderSide | false | undefined): string {
  if (!side) return 'none';
  return `1px ${side.style === 'double' ? 'double' : 'solid'} ${side.color}`;
}

/** Draw one style as a THUMB_ROWS×THUMB_COLS grid, per cell through the very
 *  cellStyleLayer the layout uses, with the table-level borders filling the
 *  sides the layer leaves open. */
function thumbnail(style: ResolvedTableStyle, look: TableLook): HTMLElement {
  const grid = document.createElement('div');
  grid.className = 'bb-tsp-thumb';
  const t: TableBorders = style.table.borders ?? {};
  for (let r = 0; r < THUMB_ROWS; r++) {
    for (let c = 0; c < THUMB_COLS; c++) {
      const layer = cellStyleLayer(style, look, {
        row: r,
        rowCount: THUMB_ROWS,
        col: c,
        colspan: 1,
        colCount: THUMB_COLS,
        header: false,
      });
      const cell = document.createElement('div');
      const b = layer.borders ?? {};
      cell.style.borderTop = edgeCss(b.top ?? (r === 0 ? t.top : t.insideH));
      cell.style.borderBottom = edgeCss(
        b.bottom ?? (r === THUMB_ROWS - 1 ? t.bottom : undefined),
      );
      cell.style.borderLeft = edgeCss(b.left ?? (c === 0 ? t.left : t.insideV));
      cell.style.borderRight = edgeCss(
        b.right ?? (c === THUMB_COLS - 1 ? t.right : undefined),
      );
      if (layer.background) cell.style.background = layer.background;
      grid.appendChild(cell);
    }
  }
  return grid;
}

export function createTableStylePanel(
  options: TableStylePanelOptions,
): TableStylePanelHandle {
  injectStyle('bb-table-style-panel', STYLE);

  const root = document.createElement('div');
  root.className = 'bb-tsp';

  const opts = document.createElement('div');
  opts.className = 'bb-tsp-opts';
  const optsLbl = document.createElement('div');
  optsLbl.className = 'bb-tsp-lbl';
  optsLbl.textContent = 'Style options';
  opts.appendChild(optsLbl);
  const checks = new Map<keyof TableLook, HTMLInputElement>();
  GATES.forEach((g, i) => {
    if (i === 3) {
      const sep = document.createElement('div');
      sep.className = 'bb-tsp-sep';
      opts.appendChild(sep);
    }
    const label = document.createElement('label');
    label.className = 'bb-tsp-chk';
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.addEventListener('change', () =>
      options.onLook({ [g.key]: input.checked } as Partial<TableLook>),
    );
    label.append(input, document.createTextNode(g.label));
    opts.appendChild(label);
    checks.set(g.key, input);
  });

  const main = document.createElement('div');
  main.className = 'bb-tsp-main';
  const mainLbl = document.createElement('div');
  mainLbl.className = 'bb-tsp-lbl';
  mainLbl.textContent = 'Table styles';
  const gallery = document.createElement('div');
  gallery.className = 'bb-tsp-gallery';
  main.append(mainLbl, gallery);

  root.append(opts, main);

  const dialog = new Dialog({
    title: 'Table style',
    modal: false,
    anchor: options.anchor,
  });
  dialog.setContent(root);

  /** What the current DOM was drawn from — refresh() skips matching input. */
  let drawnKey = '';

  const draw = (): void => {
    const current = options.current();
    const items = options.styles();
    const key = JSON.stringify([current, items.map((i) => i.id)]);
    if (key === drawnKey) return;
    drawnKey = key;

    const look = current?.look ?? null;
    for (const [gate, input] of checks) {
      input.checked = look?.[gate] ?? false;
      input.disabled = !current;
      input.closest('label')?.classList.toggle('is-disabled', !current);
    }
    gallery.replaceChildren();
    if (!current) {
      const empty = document.createElement('div');
      empty.className = 'bb-tsp-empty';
      empty.textContent = 'Click inside a table to style it.';
      gallery.appendChild(empty);
      return;
    }
    const noneTile = document.createElement('button');
    noneTile.type = 'button';
    noneTile.className =
      'bb-tsp-tile' + (current.styleId === null ? ' is-selected' : '');
    noneTile.title = 'None';
    const noneInner = document.createElement('span');
    noneInner.className = 'bb-tsp-none';
    noneInner.textContent = 'None';
    noneTile.appendChild(noneInner);
    noneTile.addEventListener('click', () => options.onPick(null));
    gallery.appendChild(noneTile);
    for (const item of items) {
      const tile = document.createElement('button');
      tile.type = 'button';
      tile.className =
        'bb-tsp-tile' + (current.styleId === item.id ? ' is-selected' : '');
      tile.title = item.name;
      tile.appendChild(thumbnail(item.style, current.look));
      tile.addEventListener('click', () => options.onPick(item));
      gallery.appendChild(tile);
    }
  };

  return {
    open() {
      drawnKey = '';
      draw();
      dialog.open();
    },
    close() {
      dialog.close();
    },
    refresh() {
      if (dialog.isOpen) draw();
    },
    destroy() {
      dialog.destroy();
    },
    get isOpen() {
      return dialog.isOpen;
    },
  };
}
