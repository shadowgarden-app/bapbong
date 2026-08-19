import { Dialog } from './dialog.js';
import { injectStyle, rovingGrid } from './internal.js';
import {
  RECENT_MAX,
  SPECIAL_CHARACTERS,
  SYMBOL_GROUPS,
  SYMBOL_NAMES,
  codePointLabel,
  parseCodePoint,
  pushRecent,
  searchSymbols,
  type SymbolEntry,
} from './symbol-sets.js';

/**
 * Insert › Symbol… — Word's dialog, in the editor's own idiom.
 *
 * NON-modal, like Word's: it floats beside the page, the caret stays where it
 * is, and the user inserts as many characters as they like before closing.
 * Two tabs — a grid of curated Unicode symbols by group (Word's is by FONT;
 * this program only ever inserts real Unicode, see symbol-sets.ts) with a
 * recently-used row, a name search and a character-code box; and Word's
 * "Special Characters" list. Insertion goes through `onInsert` — the host
 * runs the `insertText` command and returns focus to the editor — so the
 * dialog knows nothing about editor state.
 *
 * Recents are the host's to keep (a preference, per app): they arrive in
 * `recent` and every change is reported through `onRecentChange`.
 */
export interface SymbolDialogOptions {
  /** Most-recent-first row to show; usually what the host last persisted. */
  recent?: readonly string[];
  /** Called with the character to insert. */
  onInsert: (char: string) => void;
  /** Called with the new recents row after every insertion. */
  onRecentChange?: (recent: string[]) => void;
  /** Where to float, if the host has an anchor (else the dialog's default). */
  anchor?: () => DOMRect | null;
}

export interface SymbolDialogHandle {
  open(): void;
  close(): void;
  destroy(): void;
  readonly isOpen: boolean;
}

const STYLE = `
.bb-sd{display:flex;flex-direction:column;gap:12px;width:412px;max-width:100%;color:var(--bb-ui-fg,#2c2c2a)}
.bb-sd *{box-sizing:border-box}
.bb-sd-tabs{display:inline-flex;border:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-radius:6px;overflow:hidden;align-self:flex-start}
.bb-sd-tab{height:30px;padding:0 18px;border:0;border-right:0.5px solid var(--bb-ui-border,#e3e3e0);background:transparent;color:inherit;opacity:.65;font:inherit;font-size:13px;cursor:pointer}
.bb-sd-tab:last-child{border-right:0}
.bb-sd-tab[aria-selected="true"]{background:var(--bb-ui-active-bg,#e6f1fb);color:var(--bb-ui-active-fg,#0c447c);opacity:1}
.bb-sd-pane{display:flex;flex-direction:column;gap:12px}
.bb-sd-pane[hidden]{display:none}
.bb-sd-lbl{font-size:12px;opacity:.7;margin-bottom:5px}
.bb-sd-row{display:flex;gap:10px;align-items:flex-end}
.bb-sd-ctl{width:100%;height:32px;border:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-radius:6px;background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff));color:inherit;font:inherit;font-size:13px;padding:0 8px}
.bb-sd-grid{display:grid;grid-template-columns:repeat(12,1fr);gap:2px;border:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-radius:8px;padding:6px;background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff));max-height:196px;overflow:auto}
.bb-sd-grid.recent{grid-template-columns:repeat(16,1fr);max-height:none;border-color:transparent;padding:0;background:transparent;min-height:26px}
.bb-sd-cell{aspect-ratio:1;display:grid;place-items:center;border:0;border-radius:5px;background:transparent;color:inherit;font-size:18px;line-height:1;cursor:pointer;padding:0;font-family:inherit,"Noto Sans Symbols 2"}
.bb-sd-grid.recent .bb-sd-cell{border:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff));font-size:15px}
.bb-sd-cell:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-sd-cell:focus{outline:none}
.bb-sd-cell[aria-selected="true"]{background:var(--bb-ui-active-bg,#e6f1fb);box-shadow:inset 0 0 0 1.5px var(--bb-ui-active-border,#7fb2ec);color:var(--bb-ui-active-fg,#0c447c)}
.bb-sd-empty{grid-column:1/-1;font-size:12px;opacity:.55;padding:6px 2px}
.bb-sd-meta{display:grid;grid-template-columns:minmax(0,1fr) 140px;gap:10px;align-items:end}
.bb-sd-info{min-width:0;border:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-radius:8px;padding:8px 10px;display:flex;gap:12px;align-items:center;min-height:56px}
.bb-sd-big{font-size:30px;line-height:1;width:44px;text-align:center;font-family:inherit,"Noto Sans Symbols 2"}
.bb-sd-name{font-size:13px}
.bb-sd-cp{font-size:11px;opacity:.55;font-family:ui-monospace,Menlo,Consolas,monospace}
.bb-sd-code{font-family:ui-monospace,Menlo,Consolas,monospace}
.bb-sd-code[aria-invalid="true"]{border-color:#c0392b}
.bb-sd-list{border:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-radius:8px;max-height:280px;overflow:auto}
.bb-sd-list table{width:100%;border-collapse:collapse;font-size:13px}
.bb-sd-list td{padding:6px 8px;border-bottom:0.5px solid var(--bb-ui-border,#e3e3e0)}
.bb-sd-list tr:last-child td{border-bottom:0}
.bb-sd-list tr{cursor:pointer}
.bb-sd-list tr:hover td{background:var(--bb-ui-hover,#f1efe8)}
.bb-sd-list tr[aria-selected="true"] td{background:var(--bb-ui-active-bg,#e6f1fb);color:var(--bb-ui-active-fg,#0c447c)}
.bb-sd-list td.ch{width:36px;text-align:center;font-size:16px;font-family:inherit,"Noto Sans Symbols 2"}
.bb-sd-list td.k{width:120px;font-size:11px;opacity:.6;font-family:ui-monospace,Menlo,Consolas,monospace;text-align:right}
.bb-sd-foot{display:flex;justify-content:space-between;align-items:center;gap:8px}
.bb-sd-hint{font-size:11px;opacity:.55}
.bb-sd-btns{display:flex;gap:8px}
.bb-sd-btn{height:31px;padding:0 17px;border:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-radius:6px;background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff));color:inherit;font:inherit;font-size:13px;cursor:pointer}
.bb-sd-btn.primary{background:var(--bb-ui-active-bg,#e6f1fb);border-color:var(--bb-ui-active-border,#7fb2ec);color:var(--bb-ui-active-fg,#0c447c)}
.bb-sd-btn:disabled{opacity:.5;cursor:default}
`;

const GRID_COLS = 12;
const RECENT_COLS = 16;

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  cls?: string,
  text?: string,
): HTMLElementTagNameMap[K] {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
}

/** Visible stand-in for a character the eye cannot see (spaces, soft hyphen)
 *  in the grid/preview — the real character is what gets inserted. */
function glyphFor(char: string): string {
  const cp = char.codePointAt(0) ?? 0;
  if (char === ' ' || char === ' ') return '␣';
  if (cp === 0x2002 || cp === 0x2003 || cp === 0x2009) return '␣';
  if (cp === 0x00ad) return '-';
  if (cp === 0x2011) return '‑';
  return char;
}

/** Build the Symbol dialog. Returns a handle; the dialog is created closed. */
export function createSymbolDialog(
  options: SymbolDialogOptions,
): SymbolDialogHandle {
  injectStyle('bb-symbol-dialog', STYLE);
  let recent = [...(options.recent ?? [])].slice(0, RECENT_MAX);
  let selected: string | null = null;

  const root = el('div', 'bb-sd');

  // ── tabs ────────────────────────────────────────────────────────
  const tabs = el('div', 'bb-sd-tabs');
  tabs.setAttribute('role', 'tablist');
  const symTab = el('button', 'bb-sd-tab', 'Symbols');
  const spcTab = el('button', 'bb-sd-tab', 'Special characters');
  for (const t of [symTab, spcTab]) {
    t.type = 'button';
    t.setAttribute('role', 'tab');
    tabs.append(t);
  }
  const symPane = el('div', 'bb-sd-pane');
  const spcPane = el('div', 'bb-sd-pane');
  const showTab = (special: boolean) => {
    symTab.setAttribute('aria-selected', String(!special));
    spcTab.setAttribute('aria-selected', String(special));
    symPane.hidden = special;
    spcPane.hidden = !special;
  };
  symTab.addEventListener('click', () => showTab(false));
  spcTab.addEventListener('click', () => showTab(true));

  // ── Symbols: recent row ─────────────────────────────────────────
  const recentWrap = el('div');
  recentWrap.append(el('div', 'bb-sd-lbl', 'Recently used'));
  const recentGrid = el('div', 'bb-sd-grid recent');
  recentGrid.setAttribute('role', 'grid');
  recentGrid.setAttribute('aria-label', 'Recently used symbols');
  recentWrap.append(recentGrid);

  // ── Symbols: group + search ─────────────────────────────────────
  const controls = el('div', 'bb-sd-row');
  const groupCell = el('div');
  groupCell.style.flex = '1';
  groupCell.append(el('div', 'bb-sd-lbl', 'Group'));
  const groupSel = el('select', 'bb-sd-ctl');
  groupSel.setAttribute('aria-label', 'Group');
  for (const grp of SYMBOL_GROUPS) {
    const o = el('option', undefined, grp.label);
    o.value = grp.id;
    groupSel.append(o);
  }
  groupCell.append(groupSel);
  const searchCell = el('div');
  searchCell.style.width = '160px';
  searchCell.append(el('div', 'bb-sd-lbl', 'Search name'));
  const search = el('input', 'bb-sd-ctl');
  search.type = 'search';
  search.placeholder = 'e.g. ballot, arrow';
  search.setAttribute('aria-label', 'Search symbols by name');
  searchCell.append(search);
  controls.append(groupCell, searchCell);

  // ── Symbols: the grid ───────────────────────────────────────────
  const grid = el('div', 'bb-sd-grid');
  grid.setAttribute('role', 'grid');
  grid.setAttribute('aria-label', 'Symbols');

  // ── Symbols: preview + code ─────────────────────────────────────
  const meta = el('div', 'bb-sd-meta');
  const info = el('div', 'bb-sd-info');
  const big = el('div', 'bb-sd-big', '');
  const infoText = el('div');
  const nameEl = el('div', 'bb-sd-name', 'Select a symbol');
  const cpEl = el('div', 'bb-sd-cp', '');
  infoText.append(nameEl, cpEl);
  info.append(big, infoText);
  const codeCell = el('div');
  codeCell.append(el('div', 'bb-sd-lbl', 'Character code (hex)'));
  const code = el('input', 'bb-sd-ctl bb-sd-code');
  code.placeholder = 'e.g. 2611';
  code.setAttribute('aria-label', 'Character code, hexadecimal');
  code.spellcheck = false;
  codeCell.append(code);
  meta.append(info, codeCell);

  symPane.append(recentWrap, controls, grid, meta);

  // ── Special characters ──────────────────────────────────────────
  const list = el('div', 'bb-sd-list');
  const table = el('table');
  const tbody = el('tbody');
  const rows: HTMLTableRowElement[] = [];
  for (const sc of SPECIAL_CHARACTERS) {
    const tr = el('tr');
    tr.tabIndex = -1;
    tr.setAttribute('role', 'option');
    const ch = el('td', 'ch', glyphFor(sc.char));
    const nm = el('td', undefined, sc.name);
    const k = el('td', 'k', sc.hint);
    tr.append(ch, nm, k);
    tr.addEventListener('click', () => select(sc.char, tr));
    tr.addEventListener('dblclick', () => insert(sc.char));
    tr.addEventListener('keydown', (e) => {
      const i = rows.indexOf(tr);
      if (e.key === 'ArrowDown' && rows[i + 1]) {
        e.preventDefault();
        rows[i + 1].focus();
        select(SPECIAL_CHARACTERS[i + 1].char, rows[i + 1]);
      } else if (e.key === 'ArrowUp' && rows[i - 1]) {
        e.preventDefault();
        rows[i - 1].focus();
        select(SPECIAL_CHARACTERS[i - 1].char, rows[i - 1]);
      } else if (e.key === 'Enter') {
        e.preventDefault();
        insert(sc.char);
      }
    });
    tbody.append(tr);
    rows.push(tr);
  }
  table.append(tbody);
  list.append(table);
  list.setAttribute('role', 'listbox');
  spcPane.append(list);

  // ── footer ──────────────────────────────────────────────────────
  const foot = el('div', 'bb-sd-foot');
  const hint = el(
    'span',
    'bb-sd-hint',
    'Double-click a cell to insert · Enter inserts the selection',
  );
  const btns = el('div', 'bb-sd-btns');
  const closeBtn = el('button', 'bb-sd-btn', 'Close');
  closeBtn.type = 'button';
  const insertBtn = el('button', 'bb-sd-btn primary', 'Insert');
  insertBtn.type = 'button';
  insertBtn.disabled = true;
  btns.append(closeBtn, insertBtn);
  foot.append(hint, btns);

  root.append(tabs, symPane, spcPane, foot);

  const dialog = new Dialog({
    title: 'Symbol',
    modal: false,
    anchor: options.anchor,
    className: 'bb-symbol-dialog',
  });
  dialog.setContent(root);

  // ── behaviour ───────────────────────────────────────────────────
  let selectedEl: HTMLElement | null = null;
  function select(char: string, cell?: HTMLElement | null): void {
    selected = char;
    if (selectedEl) selectedEl.removeAttribute('aria-selected');
    selectedEl = cell ?? null;
    if (selectedEl) selectedEl.setAttribute('aria-selected', 'true');
    big.textContent = glyphFor(char);
    nameEl.textContent = SYMBOL_NAMES.get(char) ?? 'Character';
    cpEl.textContent = codePointLabel(char);
    if (document.activeElement !== code)
      code.value = codePointLabel(char).slice(2);
    code.removeAttribute('aria-invalid');
    insertBtn.disabled = false;
  }

  function insert(char: string): void {
    options.onInsert(char);
    recent = pushRecent(recent, char);
    options.onRecentChange?.([...recent]);
    renderRecent();
  }

  function cellFor(e: SymbolEntry): HTMLButtonElement {
    const b = el('button', 'bb-sd-cell', glyphFor(e.char));
    b.type = 'button';
    b.title = `${e.name} · ${codePointLabel(e.char)}`;
    b.setAttribute('aria-label', e.name);
    b.addEventListener('click', () => select(e.char, b));
    b.addEventListener('dblclick', () => insert(e.char));
    b.addEventListener('keydown', (ev) => {
      if (ev.key === 'Enter' || ev.key === ' ') {
        ev.preventDefault();
        select(e.char, b);
        insert(e.char);
      }
    });
    b.addEventListener('focus', () => select(e.char, b));
    return b;
  }

  function fillGrid(
    target: HTMLElement,
    entries: readonly SymbolEntry[],
    cols: number,
    empty: string,
  ): void {
    target.replaceChildren();
    if (entries.length === 0) {
      target.append(el('div', 'bb-sd-empty', empty));
      return;
    }
    const cells: HTMLElement[][] = [];
    entries.forEach((e, i) => {
      const b = cellFor(e);
      target.append(b);
      if (i % cols === 0) cells.push([]);
      cells[cells.length - 1].push(b);
    });
    rovingGrid(cells);
  }

  function renderRecent(): void {
    fillGrid(
      recentGrid,
      recent.map((c) => ({
        char: c,
        name: SYMBOL_NAMES.get(c) ?? 'Character',
      })),
      RECENT_COLS,
      'Nothing yet — inserted symbols appear here.',
    );
  }

  function renderGroup(): void {
    const q = search.value.trim();
    const entries = q
      ? searchSymbols(q)
      : (SYMBOL_GROUPS.find((g) => g.id === groupSel.value)?.entries ?? []);
    fillGrid(grid, entries, GRID_COLS, 'No symbol matches that name.');
    // First cell selected so Enter/Insert have something to act on.
    const first = grid.querySelector<HTMLElement>('.bb-sd-cell');
    if (first && entries[0]) select(entries[0].char, first);
  }

  groupSel.addEventListener('change', () => {
    search.value = '';
    renderGroup();
  });
  search.addEventListener('input', renderGroup);
  code.addEventListener('input', () => {
    const ch = parseCodePoint(code.value);
    if (ch) {
      code.removeAttribute('aria-invalid');
      // A cell in the current grid for it? Select that; else preview alone.
      const cell = Array.from(
        grid.querySelectorAll<HTMLElement>('.bb-sd-cell'),
      ).find((c) => c.textContent === glyphFor(ch));
      select(ch, cell ?? null);
    } else {
      code.setAttribute('aria-invalid', code.value.trim() ? 'true' : 'false');
      insertBtn.disabled = true;
      selected = null;
    }
  });
  code.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && selected) {
      e.preventDefault();
      insert(selected);
    }
  });
  insertBtn.addEventListener('click', () => {
    if (selected) insert(selected);
  });
  closeBtn.addEventListener('click', () => dialog.close());

  showTab(false);
  renderRecent();
  renderGroup();

  return {
    open() {
      dialog.open();
      // Land on the grid so arrows work at once (Word focuses the grid too).
      const first = grid.querySelector<HTMLElement>('.bb-sd-cell');
      first?.focus();
    },
    close() {
      dialog.close();
    },
    destroy() {
      dialog.destroy();
    },
    get isOpen() {
      return dialog.isOpen;
    },
  };
}
