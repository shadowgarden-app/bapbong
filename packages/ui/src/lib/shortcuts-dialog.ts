import type {
  Command,
  KeybindingRegistry,
  KeybindingRow,
} from '@shadow-garden/bapbong-contracts';
import {
  filterKeybindingRows,
  formatKey,
  keybindingRows,
} from '@shadow-garden/bapbong-contracts';
import { Dialog } from './dialog.js';
import { IS_MAC, injectStyle } from './internal.js';

/**
 * Help › Keyboard shortcuts — generated from the registries, never typed by
 * hand: every binding the editor answers to and every app-level one the host
 * registered, one row per command, sorted by its title, searchable across
 * command / key / when / source, sortable by column. What the registry
 * dispatches is exactly what this shows, so the two cannot drift.
 */
export interface ShortcutsSource {
  keybindings: KeybindingRegistry;
  /** Where the command titles come from. */
  commands: { get(name: string): Command | undefined };
}

export interface ShortcutsDialogOptions {
  /** The editor's registry first, then the host's app registry. */
  sources: ShortcutsSource[];
  title?: string;
}

const STYLE = `
.bb-ks{display:flex;flex-direction:column;gap:10px;width:600px;max-width:100%;color:var(--bb-ui-fg,#2c2c2a);font-size:12px}
.bb-ks *{box-sizing:border-box}
.bb-ks-top{display:flex;gap:8px;align-items:center}
.bb-ks-search{flex:1;height:30px;border:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-radius:6px;background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff));color:inherit;font:inherit;font-size:12px;padding:0 8px}
.bb-ks-search::placeholder{color:inherit;opacity:.45}
.bb-ks-count{font-size:11px;opacity:.6;white-space:nowrap}
/* No box around the list: a bordered, tinted scroller turns the scrollbar's
   gutter into a blank column at its edge. Flat on the dialog surface, the
   rail reads like the app's own (idle-hidden overlay on a plain ground). */
.bb-ks-wrap{overflow:auto;max-height:min(60vh,420px);margin:0 -12px;padding:0 12px;scrollbar-gutter:stable}
.bb-ks table{width:100%;border-collapse:collapse}
.bb-ks th{position:sticky;top:0;z-index:1;text-align:left;font-size:11px;font-weight:600;padding:7px 10px;border-bottom:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));cursor:pointer;user-select:none;white-space:nowrap;background-color:var(--bb-ui-bg,#fff);background-image:linear-gradient(var(--bb-ui-dialog-bg,transparent),var(--bb-ui-dialog-bg,transparent))}
/* Opaque surface UNDER the dialog's own tint: a sticky header on the glass
   alone let the rows scrolling beneath show through it. The dimming goes on
   the text, not the cell, for the same reason. */
.bb-ks th span{opacity:.75}
.bb-ks th[aria-sort]{color:var(--bb-ui-active-fg,#0c447c)}
.bb-ks th[aria-sort] span{opacity:1}
.bb-ks td{padding:7px 10px;border-bottom:0.5px solid var(--bb-ui-border,#e3e3e0);vertical-align:middle}
.bb-ks tr:last-child td{border-bottom:0}
.bb-ks tbody tr:hover td{background:var(--bb-ui-hover,#f1efe8)}
.bb-ks tbody tr:hover td:first-child{border-radius:6px 0 0 6px}
.bb-ks tbody tr:hover td:last-child{border-radius:0 6px 6px 0}
.bb-ks td.cmd{width:42%}
.bb-ks td.cmd small{display:block;font-size:10px;opacity:.5;font-family:ui-monospace,Menlo,Consolas,monospace}
.bb-ks kbd{display:inline-block;min-width:20px;padding:1px 6px;margin-right:3px;border:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-bottom-width:1.5px;border-radius:5px;background:var(--bb-ui-bg,#fff);font:inherit;font-size:11px;text-align:center}
.bb-ks .or{opacity:.5;margin:0 4px 0 1px;font-size:11px}
.bb-ks td.when{font-size:11px;opacity:.7;width:24%}
.bb-ks td.src{width:12%}
.bb-ks-chip{display:inline-block;padding:1px 7px;border-radius:9px;font-size:10px;border:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff));opacity:.85;white-space:nowrap}
.bb-ks-chip.core{background:var(--bb-ui-active-bg,#e6f1fb);border-color:var(--bb-ui-active-border,#7fb2ec);color:var(--bb-ui-active-fg,#0c447c);opacity:1}
.bb-ks-empty{padding:16px;text-align:center;font-size:12px;opacity:.55}
.bb-ks-foot{display:flex;justify-content:space-between;font-size:11px;opacity:.6}
.bb-ks mark{background:transparent;color:inherit;font-weight:600;text-decoration:underline;text-decoration-color:var(--bb-ui-active-border,#7fb2ec);text-underline-offset:2px}
`;

type SortKey = 'title' | 'keys' | 'when' | 'source';

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

/** `text` with the query's matches wrapped in <mark>, case-insensitive. */
function highlight(text: string, q: string): (string | HTMLElement)[] {
  if (!q) return [text];
  const out: (string | HTMLElement)[] = [];
  const lower = text.toLowerCase();
  let i = 0;
  for (;;) {
    const j = lower.indexOf(q, i);
    if (j < 0) break;
    if (j > i) out.push(text.slice(i, j));
    out.push(el('mark', undefined, text.slice(j, j + q.length)));
    i = j + q.length;
  }
  if (i < text.length) out.push(text.slice(i));
  return out;
}

/** Open the dialog. Modal; Esc / ✕ close and destroy it. */
export function openKeyboardShortcutsDialog(
  options: ShortcutsDialogOptions,
): void {
  injectStyle('bb-ui-shortcuts-dialog', STYLE);
  const dialog = new Dialog({
    title: options.title ?? 'Keyboard shortcuts',
    modal: true,
    wide: true,
  });
  const root = el('div', 'bb-ks');

  const top = el('div', 'bb-ks-top');
  const search = el('input', 'bb-ks-search');
  search.type = 'search';
  search.placeholder = 'Search command, key, when or source…';
  search.setAttribute('aria-label', 'Search shortcuts');
  const count = el('span', 'bb-ks-count', '');
  top.append(search, count);

  const wrap = el('div', 'bb-ks-wrap');
  const table = el('table');
  const thead = el('thead');
  const headRow = el('tr');
  const columns: { key: SortKey; label: string }[] = [
    { key: 'title', label: 'Command' },
    { key: 'keys', label: 'Keybinding' },
    { key: 'when', label: 'When' },
    { key: 'source', label: 'Source' },
  ];
  let sortKey: SortKey = 'title';
  let sortAsc = true;
  const ths = new Map<SortKey, HTMLTableCellElement>();
  for (const c of columns) {
    const th = el('th');
    th.append(el('span', undefined, c.label));
    th.addEventListener('click', () => {
      if (sortKey === c.key) sortAsc = !sortAsc;
      else {
        sortKey = c.key;
        sortAsc = true;
      }
      render();
    });
    headRow.append(th);
    ths.set(c.key, th);
  }
  thead.append(headRow);
  const tbody = el('tbody');
  table.append(thead, tbody);
  wrap.append(table);

  const foot = el('div', 'bb-ks-foot');
  const footL = el(
    'span',
    undefined,
    'Sorted by command · click a header to sort',
  );
  const footR = el('span', undefined, 'Esc closes');
  foot.append(footL, footR);

  root.append(top, wrap, foot);
  dialog.setContent(root);

  /** All rows, from every source: title from that source's commands. */
  const allRows = (): KeybindingRow[] => {
    const rows: KeybindingRow[] = [];
    for (const src of options.sources)
      rows.push(
        ...keybindingRows(src.keybindings, (c) => src.commands.get(c)?.title),
      );
    return rows;
  };

  const kbdCell = (keys: string[]): HTMLTableCellElement => {
    const td = el('td');
    keys.forEach((k, i) => {
      if (i > 0) td.append(el('span', 'or', 'or'));
      for (const part of formatKey(k, IS_MAC))
        td.append(el('kbd', undefined, part));
    });
    return td;
  };

  const sortRows = (rows: KeybindingRow[]): KeybindingRow[] => {
    const val = (r: KeybindingRow): string =>
      sortKey === 'title'
        ? r.title
        : sortKey === 'keys'
          ? r.keys.map((k) => formatKey(k, IS_MAC).join('')).join(' ')
          : sortKey === 'when'
            ? (r.when ?? '')
            : r.source;
    const sorted = [...rows].sort((a, b) =>
      val(a).localeCompare(val(b), undefined, { sensitivity: 'base' }),
    );
    return sortAsc ? sorted : sorted.reverse();
  };

  const render = (): void => {
    const q = search.value.trim().toLowerCase();
    const rows = allRows();
    const shown = sortRows(filterKeybindingRows(rows, q, IS_MAC));
    for (const [key, th] of ths) {
      if (key === sortKey)
        th.setAttribute('aria-sort', sortAsc ? 'ascending' : 'descending');
      else th.removeAttribute('aria-sort');
    }
    tbody.replaceChildren();
    if (shown.length === 0) {
      const tr = el('tr');
      const td = el('td', 'bb-ks-empty', 'No shortcut matches.');
      td.colSpan = 4;
      tr.append(td);
      tbody.append(tr);
    }
    for (const r of shown) {
      const tr = el('tr');
      const cmd = el('td', 'cmd');
      cmd.append(...highlight(r.title, q));
      const small = el('small');
      small.append(...highlight(r.command, q));
      cmd.append(small);
      const when = el('td', 'when');
      when.append(...highlight(r.when ?? '', q));
      const src = el('td', 'src');
      const chip = el(
        'span',
        'bb-ks-chip' + (r.source === 'core' ? ' core' : ''),
      );
      chip.append(...highlight(r.source, q));
      src.append(chip);
      tr.append(cmd, kbdCell(r.keys), when, src);
      tbody.append(tr);
    }
    count.textContent =
      shown.length === rows.length
        ? `${rows.length} shortcuts`
        : `${shown.length} of ${rows.length}`;
    footL.textContent = `Sorted by ${columns.find((c) => c.key === sortKey)?.label.toLowerCase()} · click a header to sort`;
  };
  search.addEventListener('input', render);

  dialog.onClose(() => dialog.destroy());
  render();
  dialog.open();
  search.focus();
}
