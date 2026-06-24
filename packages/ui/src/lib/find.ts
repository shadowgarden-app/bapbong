import { Dialog } from './dialog.js';
import { injectStyle } from './internal.js';

/** Query / match-count snapshot the find plugin emits. */
export interface FindState {
  query: string;
  count: number;
  /** 1-based index of the active match (0 when none). */
  active: number;
}

/**
 * The find/replace surface the panel binds to — the editor's `find` handle
 * (FindPlugin) satisfies it structurally, so this package never imports the
 * editor (same decoupling as {@link EditorHandle}).
 */
export interface FindHandle {
  setQuery(q: string): void;
  next(): void;
  prev(): void;
  clear(): void;
  replaceCurrent(text: string): void;
  replaceAll(text: string): void;
  /** Subscribe to query/count/active changes; returns an unsubscribe. */
  onState(cb: (s: FindState) => void): () => void;
}

export interface FindDialogOptions {
  /** Modal vs non-modal (default false — keeps the document editable). */
  modal?: boolean;
  /** Pin the panel's top-right inside this rect (e.g. the canvas viewport). */
  anchor?: () => DOMRect | null;
  /** Open on Ctrl/Cmd+F (pre-empting the browser's native find). Default true. */
  shortcut?: boolean;
  /** Override the (English) default labels for i18n. */
  labels?: Partial<{
    title: string;
    find: string;
    replace: string;
    prev: string;
    next: string;
    replaceOne: string;
    replaceAll: string;
  }>;
}

export interface FindDialogHandle {
  open(): void;
  close(): void;
  destroy(): void;
}

const DEFAULT_LABELS = {
  title: 'Find and replace',
  find: 'Find…',
  replace: 'Replace…',
  prev: 'Previous',
  next: 'Next',
  replaceOne: 'Replace',
  replaceAll: 'Replace all',
};

const STYLE = `
.bb-find{display:flex;flex-direction:column;gap:8px;min-width:300px}
.bb-find *{box-sizing:border-box}
.bb-find-row{display:flex;gap:6px;align-items:center}
.bb-find-input{flex:1 1 auto;height:30px;padding:0 9px;border:1px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;font:inherit;font-size:13px;background:var(--bb-ui-bg,#fff);color:inherit}
.bb-find-count{min-width:44px;text-align:center;font-size:12px;opacity:.65;font-variant-numeric:tabular-nums}
.bb-find-btn{height:30px;min-width:30px;padding:0 10px;border:1px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;background:var(--bb-ui-bg,#fff);color:inherit;font:inherit;font-size:13px;cursor:pointer}
.bb-find-btn:hover:not(:disabled){background:var(--bb-ui-hover,#f1efe8)}
.bb-find-btn:disabled{opacity:.4;cursor:default}
`;

/**
 * Build a find/replace panel inside a {@link Dialog}, bound to `find`
 * (`editor.find`). Returns open/close/destroy; opening focuses the query input,
 * closing clears the search so highlights go away. The host (any framework)
 * just opens it from a menu / shortcut.
 */
export function createFindDialog(find: FindHandle, options: FindDialogOptions = {}): FindDialogHandle {
  injectStyle('bb-ui-find-styles', STYLE);
  const labels = { ...DEFAULT_LABELS, ...(options.labels ?? {}) };

  const root = document.createElement('div');
  root.className = 'bb-find';

  const findInput = document.createElement('input');
  findInput.type = 'text';
  findInput.className = 'bb-find-input';
  findInput.placeholder = labels.find;
  findInput.setAttribute('aria-label', labels.find);

  const count = document.createElement('span');
  count.className = 'bb-find-count';
  count.textContent = '0';

  const mkBtn = (label: string, title: string) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'bb-find-btn';
    b.textContent = label;
    b.title = title;
    b.setAttribute('aria-label', title);
    return b;
  };
  const prev = mkBtn('◀', labels.prev);
  const next = mkBtn('▶', labels.next);

  const replaceInput = document.createElement('input');
  replaceInput.type = 'text';
  replaceInput.className = 'bb-find-input';
  replaceInput.placeholder = labels.replace;
  replaceInput.setAttribute('aria-label', labels.replace);
  const replaceOne = mkBtn(labels.replaceOne, labels.replaceOne);
  const replaceAll = mkBtn(labels.replaceAll, labels.replaceAll);

  const row1 = document.createElement('div');
  row1.className = 'bb-find-row';
  row1.append(findInput, count, prev, next);
  const row2 = document.createElement('div');
  row2.className = 'bb-find-row';
  row2.append(replaceInput, replaceOne, replaceAll);
  root.append(row1, row2);

  const matchButtons = [prev, next, replaceOne, replaceAll];
  findInput.addEventListener('input', () => find.setQuery(findInput.value));
  findInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      find.next();
    }
  });
  prev.addEventListener('click', () => find.prev());
  next.addEventListener('click', () => find.next());
  replaceOne.addEventListener('click', () => find.replaceCurrent(replaceInput.value));
  replaceAll.addEventListener('click', () => find.replaceAll(replaceInput.value));

  const off = find.onState((s) => {
    count.textContent = s.count ? `${s.active}/${s.count}` : '0';
    for (const b of matchButtons) b.disabled = s.count === 0;
  });
  for (const b of matchButtons) b.disabled = true;

  const dialog = new Dialog({
    title: labels.title,
    modal: options.modal ?? false,
    anchor: options.anchor,
    className: 'bb-find-dialog',
  });
  dialog.setContent(root);
  // Closing the panel clears the search so highlights/decorations disappear.
  dialog.onClose(() => {
    find.clear();
    findInput.value = '';
  });

  const openPanel = () => {
    dialog.open();
    findInput.focus();
    findInput.select();
  };

  // Ctrl/Cmd+F opens the panel and pre-empts the browser's native find.
  const onHotkey = (e: KeyboardEvent) => {
    if ((e.metaKey || e.ctrlKey) && !e.altKey && (e.key === 'f' || e.key === 'F')) {
      e.preventDefault();
      openPanel();
    }
  };
  if (options.shortcut !== false) document.addEventListener('keydown', onHotkey);

  return {
    open: openPanel,
    close() {
      dialog.close();
    },
    destroy() {
      document.removeEventListener('keydown', onHotkey);
      off();
      dialog.destroy();
    },
  };
}
