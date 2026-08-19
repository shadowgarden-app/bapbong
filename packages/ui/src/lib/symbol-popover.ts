import { injectStyle, placeFloating, rovingGrid } from './internal.js';
import {
  RECENT_MAX,
  SYMBOL_GROUPS,
  SYMBOL_NAMES,
  codePointLabel,
  parseCodePoint,
  pushRecent,
  searchSymbols,
  type SymbolEntry,
} from './symbol-sets.js';

/**
 * The toolbar's Ω: a compact symbol picker hanging under the button — the
 * quick path (Word's ribbon gallery of recent symbols), where the Symbol
 * panel (symbol-dialog.ts) is the full one (Word's "More Symbols…").
 *
 * Same construction as the colour picker: a modal `<dialog>` in the top
 * layer (so it opens over other panels), placed under its anchor, dismissed by
 * a click outside, Esc or a scroll. Search box (a name, or a code like
 * `2611`/`U+2611`), the recently-used row, the groups as chips, an 8-column
 * grid, and the selected symbol's name beside a "More symbols…" link. A click
 * on a cell INSERTS at once — and the picker stays open for the next one; it
 * is a picker, not a form.
 */
export interface SymbolPopoverOptions {
  /** The control the picker hangs from (the toolbar's Ω). */
  anchor: HTMLElement;
  /** Most-recent-first row to show. */
  recent?: readonly string[];
  onInsert: (char: string) => void;
  onRecentChange?: (recent: string[]) => void;
  /** "More symbols…": the host opens the full panel; the picker closes. */
  onMore?: () => void;
  /** Called after the picker closes, for the host to hand focus back. */
  onClose?: () => void;
}

const STYLE = `
.bb-spk{position:fixed;margin:0;max-width:none;max-height:none;z-index:1300;width:276px;padding:10px;background:var(--bb-ui-menu-bg,#fff);-webkit-backdrop-filter:var(--bb-ui-pop-filter,none);backdrop-filter:var(--bb-ui-pop-filter,none);color:var(--bb-ui-fg,#2c2c2a);border:1px solid var(--bb-ui-pop-border,var(--bb-ui-border,#e3e3e0));border-radius:10px;box-shadow:0 8px 28px rgba(0,0,0,.16);font-family:var(--bb-ui-font,system-ui,-apple-system,sans-serif);font-size:12px;box-sizing:border-box;display:flex;flex-direction:column;gap:8px}
.bb-spk *{box-sizing:border-box}
.bb-spk::backdrop{background:transparent}
.bb-spk-search{width:100%;height:28px;border:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-radius:6px;background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff));color:inherit;font:inherit;font-size:12px;padding:0 8px}
.bb-spk-search::placeholder{color:inherit;opacity:.45}
.bb-spk-lbl{font-size:10.5px;opacity:.65;margin:0 0 4px}
.bb-spk-recent{display:grid;grid-template-columns:repeat(8,1fr);gap:3px;min-height:26px}
.bb-spk-recent .bb-spk-cell{border:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff));font-size:14px}
.bb-spk-empty{grid-column:1/-1;font-size:11px;opacity:.5;padding:4px 2px}
.bb-spk-chips{display:flex;flex-wrap:wrap;gap:3px}
.bb-spk-chip{padding:2px 6px;border-radius:10px;border:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff));color:inherit;font:inherit;font-size:10px;opacity:.8;cursor:pointer}
.bb-spk-chip:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-spk-chip[aria-pressed="true"]{background:var(--bb-ui-active-bg,#e6f1fb);color:var(--bb-ui-active-fg,#0c447c);border-color:var(--bb-ui-active-border,#7fb2ec);opacity:1}
.bb-spk-grid{display:grid;grid-template-columns:repeat(8,1fr);gap:2px;border:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-radius:7px;padding:4px;background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff));max-height:150px;overflow:auto}
.bb-spk-cell{aspect-ratio:1;display:grid;place-items:center;border:0;border-radius:5px;background:transparent;color:inherit;font-size:16px;line-height:1;padding:0;cursor:pointer;font-family:inherit,"Noto Sans Symbols 2"}
.bb-spk-cell:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-spk-cell:focus{outline:none;box-shadow:inset 0 0 0 1.5px var(--bb-ui-active-border,#7fb2ec)}
.bb-spk-cell[aria-selected="true"]{background:var(--bb-ui-active-bg,#e6f1fb);color:var(--bb-ui-active-fg,#0c447c)}
.bb-spk-foot{display:flex;justify-content:space-between;align-items:center;gap:8px;font-size:11px}
.bb-spk-name{opacity:.75;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bb-spk-more{border:0;background:transparent;color:var(--bb-ui-active-fg,#0c447c);font:inherit;font-size:11px;cursor:pointer;padding:2px 4px;border-radius:4px;white-space:nowrap}
.bb-spk-more:hover{background:var(--bb-ui-hover,#f1efe8)}
`;

const COLS = 8;
/** The picker's recent row is ONE row — the last eight; the panel keeps all
 *  sixteen. Word's ribbon gallery is a single row too. */
const RECENT_ROW = COLS;

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

/** Open the picker under `anchor`. Returns a close handle. */
export function openSymbolPopover(options: SymbolPopoverOptions): {
  close(): void;
} {
  injectStyle('bb-ui-symbol-popover', STYLE);
  const { anchor } = options;
  let recent = [...(options.recent ?? [])].slice(0, RECENT_MAX);
  let groupId = SYMBOL_GROUPS[0].id;
  let selected: string | null = null;

  const panel = document.createElement('dialog');
  panel.className = 'bb-spk';
  panel.setAttribute('aria-label', 'Insert symbol');

  let closed = false;
  const close = (): void => {
    if (closed) return;
    closed = true;
    window.removeEventListener('scroll', close, true);
    if (panel.open) panel.close();
    panel.remove();
    options.onClose?.();
  };

  // ── search ─────────────────────────────────────────────────────
  const search = el('input', 'bb-spk-search');
  search.type = 'search';
  search.placeholder = 'Search name or code (U+2611)';
  search.setAttribute('aria-label', 'Search symbols by name or code');
  search.spellcheck = false;

  // ── recent ─────────────────────────────────────────────────────
  const recentWrap = el('div');
  recentWrap.append(el('div', 'bb-spk-lbl', 'Recent'));
  const recentGrid = el('div', 'bb-spk-recent');
  recentGrid.setAttribute('role', 'grid');
  recentGrid.setAttribute('aria-label', 'Recently used symbols');
  recentWrap.append(recentGrid);

  // ── group chips ────────────────────────────────────────────────
  const chips = el('div', 'bb-spk-chips');
  const chipEls = new Map<string, HTMLButtonElement>();
  for (const grp of SYMBOL_GROUPS) {
    const c = el('button', 'bb-spk-chip', grp.short);
    c.type = 'button';
    c.title = grp.label;
    c.addEventListener('click', () => {
      groupId = grp.id;
      search.value = '';
      render();
    });
    chips.append(c);
    chipEls.set(grp.id, c);
  }

  // ── grid ───────────────────────────────────────────────────────
  const grid = el('div', 'bb-spk-grid');
  grid.setAttribute('role', 'grid');
  grid.setAttribute('aria-label', 'Symbols');

  // ── footer ─────────────────────────────────────────────────────
  const foot = el('div', 'bb-spk-foot');
  const nameEl = el('span', 'bb-spk-name', '');
  const more = el('button', 'bb-spk-more', 'More symbols…');
  more.type = 'button';
  more.addEventListener('click', () => {
    close();
    options.onMore?.();
  });
  foot.append(nameEl, more);
  if (!options.onMore) more.hidden = true;

  panel.append(search, recentWrap, chips, grid, foot);

  // ── behaviour ──────────────────────────────────────────────────
  let selectedEl: HTMLElement | null = null;
  const select = (char: string, cell: HTMLElement | null): void => {
    selected = char;
    selectedEl?.removeAttribute('aria-selected');
    selectedEl = cell;
    selectedEl?.setAttribute('aria-selected', 'true');
    nameEl.textContent = `${SYMBOL_NAMES.get(char) ?? 'Character'} · ${codePointLabel(char)}`;
  };
  const insert = (char: string): void => {
    const active = document.activeElement;
    options.onInsert(char);
    recent = pushRecent(recent, char);
    options.onRecentChange?.([...recent]);
    renderRecent();
    // Stay usable: the host's insert ends by focusing the editor, which a
    // modal dialog's focus containment turns into a no-op — but be explicit,
    // and land somewhere alive if the recent row was just rebuilt under us.
    queueMicrotask(() => {
      if (active instanceof HTMLElement && active.isConnected) active.focus();
      else search.focus();
    });
  };
  const cellFor = (e: SymbolEntry): HTMLButtonElement => {
    const b = el('button', 'bb-spk-cell', e.char);
    b.type = 'button';
    b.title = `${e.name} · ${codePointLabel(e.char)}`;
    b.setAttribute('aria-label', e.name);
    b.addEventListener('focus', () => select(e.char, b));
    b.addEventListener('click', () => {
      select(e.char, b);
      insert(e.char);
    });
    b.addEventListener('keydown', (ev) => {
      if (ev.key === 'Enter' || ev.key === ' ') {
        ev.preventDefault();
        select(e.char, b);
        insert(e.char);
      }
    });
    return b;
  };
  const fill = (
    target: HTMLElement,
    entries: readonly SymbolEntry[],
    empty: string,
  ): void => {
    target.replaceChildren();
    if (entries.length === 0) {
      target.append(el('div', 'bb-spk-empty', empty));
      return;
    }
    const cells: HTMLElement[][] = [];
    entries.forEach((e, i) => {
      const b = cellFor(e);
      target.append(b);
      if (i % COLS === 0) cells.push([]);
      cells[cells.length - 1].push(b);
    });
    rovingGrid(cells);
  };
  const renderRecent = (): void =>
    fill(
      recentGrid,
      recent.slice(0, RECENT_ROW).map((c) => ({
        char: c,
        name: SYMBOL_NAMES.get(c) ?? 'Character',
      })),
      'Nothing yet.',
    );
  const render = (): void => {
    const q = search.value.trim();
    const byCode = q ? parseCodePoint(q) : null;
    const entries: SymbolEntry[] = q
      ? byCode && !searchSymbols(q).length
        ? [{ char: byCode, name: SYMBOL_NAMES.get(byCode) ?? 'Character' }]
        : searchSymbols(q)
      : (SYMBOL_GROUPS.find((g) => g.id === groupId)?.entries ?? []);
    for (const [id, c] of chipEls)
      c.setAttribute('aria-pressed', String(!q && id === groupId));
    fill(grid, entries, 'No match.');
    const first = grid.querySelector<HTMLElement>('.bb-spk-cell');
    if (first && entries[0]) select(entries[0].char, first);
    else {
      selected = null;
      nameEl.textContent = '';
    }
  };
  search.addEventListener('input', render);
  search.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && selected) {
      e.preventDefault();
      insert(selected);
    } else if (e.key === 'ArrowDown') {
      e.preventDefault();
      grid.querySelector<HTMLElement>('.bb-spk-cell')?.focus();
    }
  });

  // ── mount + dismiss (as the colour picker does) ────────────────
  document.body.append(panel);
  panel.showModal();
  const r = anchor.getBoundingClientRect();
  placeFloating(panel, { x: r.left, y: r.top, height: r.height });
  panel.addEventListener('close', close);
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

  renderRecent();
  render();
  search.focus();
  return { close };
}
