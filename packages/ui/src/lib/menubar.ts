import type { Collection, Command } from '@shadow-garden/bapbong-contracts';
import { type EditorHandle, type EditorStateOf, injectStyle } from './internal.js';

/** A row in a dropdown: run `command` (and read its active state), or a rule. */
export interface MenuItem {
  command: string;
  /** Row label; defaults to a built-in title or the command name. */
  label?: string;
}
export type MenuEntry = MenuItem | 'separator';

/** A top-level menu (its title + dropdown entries). */
export interface Menu {
  label: string;
  entries: MenuEntry[];
}

export interface MenubarOptions {
  /** Menus to render. Default: a single "Format" menu (marks, then a separator,
   *  then alignments) derived from the registry. */
  menus?: Menu[];
  /** Command-name → row label override, merged over the built-in labels. */
  labels?: Record<string, string>;
}

export interface MenubarHandle {
  destroy(): void;
}

const DEFAULT_LABELS: Record<string, string> = {
  bold: 'Bold',
  italic: 'Italic',
  underline: 'Underline',
  strike: 'Strikethrough',
  'align-left': 'Align left',
  'align-center': 'Center',
  'align-right': 'Align right',
  'align-justify': 'Justify',
};

const STYLE = `
.bb-menubar{display:flex;gap:2px;align-items:center;padding:2px 6px;font-family:var(--bb-ui-font,system-ui,-apple-system,sans-serif);color:var(--bb-ui-fg,#2c2c2a);background:var(--bb-ui-bg,#fff);border-bottom:1px solid var(--bb-ui-border,#e3e3e0);box-sizing:border-box}
.bb-menubar *{box-sizing:border-box}
.bb-menubar-menu{position:relative}
.bb-menubar-title{height:28px;padding:0 10px;border:0;border-radius:6px;background:transparent;color:inherit;font:inherit;font-size:13px;cursor:pointer}
.bb-menubar-title:hover,.bb-menubar-title[aria-expanded="true"]{background:var(--bb-ui-hover,#f1efe8)}
.bb-menu{position:absolute;top:100%;left:0;min-width:184px;margin-top:3px;padding:4px;background:var(--bb-ui-menu-bg,#fff);border:1px solid var(--bb-ui-border,#e3e3e0);border-radius:8px;box-shadow:0 8px 28px rgba(0,0,0,.14);z-index:1000}
.bb-menu[hidden]{display:none}
.bb-menu-item{display:flex;align-items:center;gap:8px;width:100%;height:30px;padding:0 12px 0 6px;border:0;border-radius:5px;background:transparent;color:inherit;font:inherit;font-size:13px;text-align:left;white-space:nowrap;cursor:pointer}
.bb-menu-item:hover,.bb-menu-item:focus{background:var(--bb-ui-hover,#f1efe8);outline:none}
.bb-menu-check{width:14px;flex:none;display:inline-flex;justify-content:center;color:var(--bb-ui-active-fg,#0c447c)}
.bb-menu-sep{height:1px;margin:4px 6px;background:var(--bb-ui-border,#e3e3e0)}
`;

/** Default menus: a "Format" menu of marks, a separator, then alignments —
 *  only the commands the registry actually has. */
export function defaultMenus(commands: Collection<Command>): Menu[] {
  const names = [...commands].map((c) => c.name);
  const marks = ['bold', 'italic', 'underline', 'strike'].filter((n) => commands.has(n));
  const aligns = names.filter((n) => n.startsWith('align-'));
  const entries: MenuEntry[] = [
    ...marks.map((command) => ({ command })),
    ...(marks.length && aligns.length ? (['separator'] as MenuEntry[]) : []),
    ...aligns.map((command) => ({ command })),
  ];
  return entries.length ? [{ label: 'Format', entries }] : [];
}

/**
 * Render a menubar into `host` and wire it to `editor`. Top-level titles open
 * dropdowns of command rows (with a check for active toggles). The lib owns the
 * DOM, styling, open/close, keyboard nav and active state; the host framework
 * only supplies the element + the editor handle, then calls `destroy()`.
 */
export function mountMenubar(
  host: HTMLElement,
  editor: EditorHandle,
  options: MenubarOptions = {},
): MenubarHandle {
  injectStyle('bb-ui-menubar-styles', STYLE);
  const labels = { ...DEFAULT_LABELS, ...(options.labels ?? {}) };
  const menus = options.menus ?? defaultMenus(editor.commands);

  const root = document.createElement('div');
  root.className = 'bb-menubar';
  root.setAttribute('role', 'menubar');

  // Every command row across all menus, for active-state refresh.
  const rows: Array<{ name: string; checkEl: HTMLElement }> = [];
  // Each menu's title + dropdown, for open/close coordination.
  const panels: Array<{ title: HTMLButtonElement; dropdown: HTMLElement }> = [];
  let openIdx = -1;
  let latest: EditorStateOf | null = null;

  const itemsOf = (dd: HTMLElement) =>
    Array.from(dd.querySelectorAll<HTMLButtonElement>('.bb-menu-item'));

  const close = (refocusTitle = false): void => {
    if (openIdx < 0) return;
    const p = panels[openIdx];
    p.dropdown.hidden = true;
    p.title.setAttribute('aria-expanded', 'false');
    if (refocusTitle) p.title.focus();
    openIdx = -1;
  };

  const open = (idx: number, focusFirst = false): void => {
    if (openIdx === idx) return;
    close();
    const p = panels[idx];
    p.dropdown.hidden = false;
    p.title.setAttribute('aria-expanded', 'true');
    openIdx = idx;
    if (latest) refresh(latest);
    if (focusFirst) itemsOf(p.dropdown)[0]?.focus();
  };

  menus.forEach((menu, idx) => {
    const wrap = document.createElement('div');
    wrap.className = 'bb-menubar-menu';

    const title = document.createElement('button');
    title.type = 'button';
    title.className = 'bb-menubar-title';
    title.textContent = menu.label;
    title.setAttribute('role', 'menuitem');
    title.setAttribute('aria-haspopup', 'true');
    title.setAttribute('aria-expanded', 'false');

    const dropdown = document.createElement('div');
    dropdown.className = 'bb-menu';
    dropdown.setAttribute('role', 'menu');
    dropdown.hidden = true;

    title.addEventListener('mousedown', (e) => e.preventDefault()); // keep editor selection
    title.addEventListener('click', () => (openIdx === idx ? close() : open(idx)));
    // Hovering another title while a menu is open switches to it.
    title.addEventListener('mouseenter', () => {
      if (openIdx >= 0 && openIdx !== idx) open(idx);
    });
    title.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowDown' || e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        open(idx, true);
      }
    });

    for (const entry of menu.entries) {
      if (entry === 'separator') {
        const sep = document.createElement('div');
        sep.className = 'bb-menu-sep';
        sep.setAttribute('role', 'separator');
        dropdown.appendChild(sep);
        continue;
      }
      const cmd = editor.commands.get(entry.command);
      if (!cmd) continue; // skip commands the schema/registry doesn't provide
      const item = document.createElement('button');
      item.type = 'button';
      item.className = 'bb-menu-item';
      item.setAttribute('role', 'menuitemcheckbox');
      const check = document.createElement('span');
      check.className = 'bb-menu-check';
      const text = document.createElement('span');
      text.textContent = entry.label ?? labels[entry.command] ?? entry.command;
      item.append(check, text);
      item.addEventListener('mousedown', (e) => e.preventDefault());
      item.addEventListener('click', () => {
        if (latest) editor.commands.get(entry.command)?.run(latest, (tr) => editor.dispatch(tr));
        close();
        editor.focus();
      });
      item.addEventListener('keydown', (e) => onItemKey(e, dropdown));
      dropdown.appendChild(item);
      rows.push({ name: entry.command, checkEl: check });
    }

    wrap.append(title, dropdown);
    root.appendChild(wrap);
    panels.push({ title, dropdown });
  });

  function onItemKey(e: KeyboardEvent, dd: HTMLElement): void {
    const items = itemsOf(dd);
    const i = items.indexOf(document.activeElement as HTMLButtonElement);
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      items[(i + 1) % items.length]?.focus();
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      items[(i - 1 + items.length) % items.length]?.focus();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      close(true);
    }
  }

  host.appendChild(root);

  // Close when focus/click leaves the menubar.
  const onDocPointer = (e: Event): void => {
    if (openIdx >= 0 && !root.contains(e.target as Node)) close();
  };
  document.addEventListener('pointerdown', onDocPointer);

  function refresh(state: EditorStateOf): void {
    latest = state;
    if (openIdx < 0) return; // checks only matter while a menu is visible
    for (const { name, checkEl } of rows) {
      const active = editor.commands.get(name)?.isActive?.(state) ?? false;
      checkEl.textContent = active ? '✓' : '';
    }
  }

  const off = editor.onChange((c) => refresh(c.state));
  try {
    latest = editor.state;
  } catch {
    /* no document yet */
  }

  return {
    destroy() {
      off();
      document.removeEventListener('pointerdown', onDocPointer);
      root.remove();
    },
  };
}
