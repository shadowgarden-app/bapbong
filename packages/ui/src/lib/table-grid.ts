import { injectStyle } from './internal.js';

export interface TableGridOptions {
  /** Max selectable rows / columns (the grid size). Default 8 × 10. */
  maxRows?: number;
  maxCols?: number;
  /** Called with the 1-based size when a cell is clicked. */
  onPick: (rows: number, cols: number) => void;
}

const STYLE = `
.bb-grid{display:flex;flex-direction:column;gap:6px;user-select:none}
.bb-grid-cells{display:grid;gap:2px}
.bb-grid-cell{width:15px;height:15px;border:1px solid var(--bb-ui-border,#d8d6cf);border-radius:2px;background:var(--bb-ui-bg,#fff);cursor:pointer;padding:0}
.bb-grid-cell.on{background:var(--bb-ui-active-bg,#e6f1fb);border-color:var(--bb-ui-active-border,#7fb2ec)}
.bb-grid-label{font-size:12px;text-align:center;opacity:.7}
`;

/**
 * A Word/Docs-style table size picker: hover the grid to size, click to choose.
 * Returns a detached element to drop into a menu widget flyout; calls
 * `onPick(rows, cols)` (1-based) on click.
 */
export function tableGridPicker(options: TableGridOptions): HTMLElement {
  injectStyle('bb-ui-grid-styles', STYLE);
  const maxRows = options.maxRows ?? 8;
  const maxCols = options.maxCols ?? 10;

  const root = document.createElement('div');
  root.className = 'bb-grid';
  const cellsEl = document.createElement('div');
  cellsEl.className = 'bb-grid-cells';
  cellsEl.style.gridTemplateColumns = `repeat(${maxCols}, 15px)`;
  const label = document.createElement('div');
  label.className = 'bb-grid-label';
  label.textContent = '0 × 0';

  const cells: HTMLButtonElement[] = [];
  const highlight = (r: number, c: number): void => {
    for (let i = 0; i < cells.length; i++) {
      const row = Math.floor(i / maxCols);
      const col = i % maxCols;
      cells[i].classList.toggle('on', row <= r && col <= c);
    }
    label.textContent = `${r + 1} × ${c + 1}`;
  };

  for (let r = 0; r < maxRows; r++) {
    for (let c = 0; c < maxCols; c++) {
      const cell = document.createElement('button');
      cell.type = 'button';
      cell.className = 'bb-grid-cell';
      cell.addEventListener('mousedown', (e) => e.preventDefault());
      cell.addEventListener('mouseenter', () => highlight(r, c));
      cell.addEventListener('click', () => options.onPick(r + 1, c + 1));
      cells.push(cell);
      cellsEl.appendChild(cell);
    }
  }

  root.append(cellsEl, label);
  return root;
}
