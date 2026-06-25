import { injectStyle } from './internal.js';

/** A one-shot context-menu row. */
export interface ContextMenuItem {
  label: string;
  run: () => void;
  /** Render disabled when false. */
  enabled?: boolean;
  /** Optional right-aligned hint (e.g. '⌘X'). */
  shortcut?: string;
}
export type ContextMenuEntry = ContextMenuItem | 'separator';

export interface ContextMenuHandle {
  close(): void;
}

const STYLE = `
.bb-ctx{position:fixed;z-index:1200;min-width:208px;padding:4px;background:var(--bb-ui-menu-bg,#fff);color:var(--bb-ui-fg,#2c2c2a);border:1px solid var(--bb-ui-border,#e3e3e0);border-radius:8px;box-shadow:0 8px 28px rgba(0,0,0,.16);font-family:var(--bb-ui-font,system-ui,-apple-system,sans-serif)}
.bb-ctx *{box-sizing:border-box}
.bb-ctx-item{display:flex;align-items:center;gap:16px;width:100%;height:30px;padding:0 10px;border:0;border-radius:5px;background:transparent;color:inherit;font:inherit;font-size:13px;text-align:left;white-space:nowrap;cursor:pointer}
.bb-ctx-item:hover:not(:disabled),.bb-ctx-item:focus{background:var(--bb-ui-hover,#f1efe8);outline:none}
.bb-ctx-item:disabled{opacity:.4;cursor:default}
.bb-ctx-label{flex:1 1 auto}
.bb-ctx-shortcut{flex:none;opacity:.5;font-size:12px}
.bb-ctx-sep{height:1px;margin:4px 6px;background:var(--bb-ui-border,#e3e3e0)}
`;

let current: (() => void) | null = null;

function closeCurrent(): void {
  if (current) {
    const dispose = current;
    current = null;
    dispose();
  }
}

/**
 * Show a context menu at viewport point `at`. One-shot rows (label + run, with
 * optional enabled/shortcut) and separators; closes on select, outside click,
 * Escape, or scroll. Only one is open at a time. Reusable for any right-click
 * surface — the editor's pointer hook gives the host the position + target.
 */
export function showContextMenu(entries: ContextMenuEntry[], at: { x: number; y: number }): ContextMenuHandle {
  injectStyle('bb-ui-contextmenu-styles', STYLE);
  closeCurrent();

  const el = document.createElement('div');
  el.className = 'bb-ctx';
  el.setAttribute('role', 'menu');

  for (const entry of entries) {
    if (entry === 'separator') {
      const sep = document.createElement('div');
      sep.className = 'bb-ctx-sep';
      sep.setAttribute('role', 'separator');
      el.appendChild(sep);
      continue;
    }
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'bb-ctx-item';
    item.setAttribute('role', 'menuitem');
    item.disabled = entry.enabled === false;
    const label = document.createElement('span');
    label.className = 'bb-ctx-label';
    label.textContent = entry.label;
    item.appendChild(label);
    if (entry.shortcut) {
      const sc = document.createElement('span');
      sc.className = 'bb-ctx-shortcut';
      sc.textContent = entry.shortcut;
      item.appendChild(sc);
    }
    item.addEventListener('mousedown', (e) => e.preventDefault());
    item.addEventListener('click', () => {
      closeCurrent();
      entry.run();
    });
    el.appendChild(item);
  }
  document.body.appendChild(el);

  // Clamp into the viewport (flip when overflowing right/bottom).
  const r = el.getBoundingClientRect();
  el.style.left = `${Math.max(4, Math.min(at.x, window.innerWidth - r.width - 4))}px`;
  el.style.top = `${Math.max(4, Math.min(at.y, window.innerHeight - r.height - 4))}px`;

  const items = () => Array.from(el.querySelectorAll<HTMLButtonElement>('.bb-ctx-item:not(:disabled)'));
  const onDown = (e: Event) => {
    if (!el.contains(e.target as Node)) closeCurrent();
  };
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape') return closeCurrent();
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      const list = items();
      const i = list.indexOf(document.activeElement as HTMLButtonElement);
      const next = e.key === 'ArrowDown' ? (i + 1) % list.length : (i - 1 + list.length) % list.length;
      list[next]?.focus();
    }
  };
  const onScroll = () => closeCurrent();
  // Defer so the opening (right-click) event doesn't immediately dismiss it.
  setTimeout(() => {
    document.addEventListener('pointerdown', onDown, true);
    document.addEventListener('keydown', onKey, true);
    window.addEventListener('scroll', onScroll, true);
  }, 0);

  current = () => {
    document.removeEventListener('pointerdown', onDown, true);
    document.removeEventListener('keydown', onKey, true);
    window.removeEventListener('scroll', onScroll, true);
    el.remove();
  };
  items()[0]?.focus();
  return { close: closeCurrent };
}
