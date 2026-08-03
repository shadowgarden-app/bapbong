import type { Collection, Command } from '@shadow-garden/bapbong-contracts';
import { type EditorHandle, type EditorStateOf, injectStyle } from './internal.js';

/** Run a registry command (and read its active/enabled state). */
export interface CommandEntry {
  command: string;
  label?: string;
}
/** Run a host-supplied action (File ▸ Open, View ▸ comment mode, Help…). */
export interface ActionEntry {
  label: string;
  run: () => void;
  isActive?: () => boolean;
  isEnabled?: () => boolean;
  shortcut?: string;
}
/** A nested dropdown. */
export interface SubmenuEntry {
  label: string;
  submenu: MenuEntry[];
}
/** Custom flyout content (e.g. a table size grid). `close` dismisses the menu. */
export interface WidgetEntry {
  label: string;
  widget: (close: () => void) => HTMLElement;
}
export type MenuEntry = 'separator' | CommandEntry | ActionEntry | SubmenuEntry | WidgetEntry;

/** A top-level menu (its title + dropdown entries). */
export interface Menu {
  label: string;
  entries: MenuEntry[];
}

export interface MenubarOptions {
  /** Menus to render. Default: a single "Format" menu derived from the registry. */
  menus?: Menu[];
  /** Command-name → row label override, merged over the built-in labels. */
  labels?: Record<string, string>;
  /** `vertical` stacks the menu titles in a column and opens each dropdown to
   *  the right, bottom-aligned so it grows upward — for a menubar docked in a
   *  corner (e.g. bottom-left). Default `horizontal` (a classic top bar). */
  orientation?: 'horizontal' | 'vertical';
}

export interface MenubarHandle {
  destroy(): void;
}

const DEFAULT_LABELS: Record<string, string> = {
  bold: 'Bold',
  italic: 'Italic',
  underline: 'Underline',
  strike: 'Strikethrough',
  superscript: 'Superscript',
  subscript: 'Subscript',
  'align-left': 'Align left',
  'align-center': 'Center',
  'align-right': 'Align right',
  'align-justify': 'Justify',
  'bullet-list': 'Bullet list',
  'ordered-list': 'Numbered list',
  'heading-1': 'Heading 1',
  'heading-2': 'Heading 2',
  'heading-3': 'Heading 3',
  'heading-4': 'Heading 4',
  'heading-5': 'Heading 5',
  'heading-6': 'Heading 6',
  undo: 'Undo',
  redo: 'Redo',
  'page-break': 'Page break',
};

const STYLE = `
.bb-menubar{display:flex;gap:2px;align-items:center;padding:2px 6px;font-family:var(--bb-ui-font,system-ui,-apple-system,sans-serif);color:var(--bb-ui-fg,#2c2c2a);background:var(--bb-ui-bg,#fff);border-bottom:1px solid var(--bb-ui-border,#e3e3e0);box-sizing:border-box}
.bb-menubar *{box-sizing:border-box}
.bb-menubar-menu{position:relative}
.bb-menubar-title{height:28px;padding:0 10px;border:0;border-radius:6px;background:transparent;color:inherit;font:inherit;font-size:13px;cursor:pointer}
.bb-menubar-title:hover,.bb-menubar-title[aria-expanded="true"]{background:var(--bb-ui-hover,#f1efe8)}
.bb-menu{position:absolute;top:100%;left:0;min-width:200px;margin-top:3px;padding:4px;background:var(--bb-ui-menu-bg,#fff);-webkit-backdrop-filter:var(--bb-ui-pop-filter,none);backdrop-filter:var(--bb-ui-pop-filter,none);border:1px solid var(--bb-ui-pop-border,var(--bb-ui-border,#e3e3e0));border-radius:8px;box-shadow:0 8px 28px rgba(0,0,0,.14);z-index:1000}
.bb-menu[hidden]{display:none}
.bb-menu-sub{position:relative}
.bb-submenu{top:-5px;left:100%;margin-left:2px;display:none}
.bb-menu-sub:hover>.bb-submenu,.bb-menu-sub:focus-within>.bb-submenu{display:block}
.bb-submenu-widget{padding:8px;min-width:0}
.bb-menu-item{display:flex;align-items:center;gap:8px;width:100%;height:30px;padding:0 10px 0 6px;border:0;border-radius:5px;background:transparent;color:inherit;font:inherit;font-size:13px;text-align:left;white-space:nowrap;cursor:pointer}
.bb-menu-item:hover,.bb-menu-item:focus{background:var(--bb-ui-hover,#f1efe8);outline:none}
.bb-menu-item:disabled{opacity:.4;cursor:default}
.bb-menu-check{width:14px;flex:none;display:inline-flex;justify-content:center;color:var(--bb-ui-active-fg,#0c447c)}
.bb-menu-label{flex:1 1 auto}
.bb-menu-shortcut{flex:none;opacity:.5;font-size:12px;padding-left:24px}
.bb-menu-arrow{flex:none;opacity:.55;padding-left:12px}
.bb-menu-sep{height:1px;margin:4px 6px;background:var(--bb-ui-border,#e3e3e0)}
.bb-menubar-v{flex-direction:column;align-items:stretch;gap:1px;border-bottom:none;padding:4px}
.bb-menubar-v .bb-menubar-title{text-align:left;width:100%}
.bb-menubar-v .bb-menu{top:auto;bottom:0;left:100%;margin-top:0;margin-left:4px}
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
 * dropdowns; entries can be registry commands (with an active check), host
 * actions, nested submenus, or custom widgets. The lib owns the DOM, styling,
 * open/close and active state; the host framework only supplies the element +
 * the editor handle, then calls `destroy()`.
 */
export function mountMenubar(
  host: HTMLElement,
  editor: EditorHandle,
  options: MenubarOptions = {},
): MenubarHandle {
  injectStyle('bb-ui-menubar-styles', STYLE);
  const labels = { ...DEFAULT_LABELS, ...(options.labels ?? {}) };
  const menus = options.menus ?? defaultMenus(editor.commands);

  const vertical = options.orientation === 'vertical';
  const root = document.createElement('div');
  root.className = 'bb-menubar' + (vertical ? ' bb-menubar-v' : '');
  root.setAttribute('role', 'menubar');
  if (vertical) root.setAttribute('aria-orientation', 'vertical');

  const panels: Array<{ title: HTMLButtonElement; dropdown: HTMLElement }> = [];
  // Rows whose check / disabled state tracks editor state (any nesting depth).
  const checks: Array<{ el: HTMLElement; active: (s: EditorStateOf) => boolean }> = [];
  const enables: Array<{ el: HTMLButtonElement; enabled: (s: EditorStateOf) => boolean }> = [];
  let openIdx = -1;
  let latest: EditorStateOf | null = null;

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
    if (focusFirst) p.dropdown.querySelector<HTMLButtonElement>('.bb-menu-item')?.focus();
  };

  const makeRow = (label: string, opts: { shortcut?: string; hasSub?: boolean }) => {
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'bb-menu-item' + (opts.hasSub ? ' bb-has-sub' : '');
    item.addEventListener('mousedown', (e) => e.preventDefault()); // keep editor selection
    const check = document.createElement('span');
    check.className = 'bb-menu-check';
    const text = document.createElement('span');
    text.className = 'bb-menu-label';
    text.textContent = label;
    item.append(check, text);
    if (opts.shortcut) {
      const sc = document.createElement('span');
      sc.className = 'bb-menu-shortcut';
      sc.textContent = opts.shortcut;
      item.appendChild(sc);
    }
    if (opts.hasSub) {
      const arrow = document.createElement('span');
      arrow.className = 'bb-menu-arrow';
      arrow.textContent = '›';
      item.appendChild(arrow);
    }
    return { item, check };
  };

  const buildEntries = (entries: MenuEntry[], container: HTMLElement): void => {
    for (const entry of entries) {
      if (entry === 'separator') {
        const sep = document.createElement('div');
        sep.className = 'bb-menu-sep';
        sep.setAttribute('role', 'separator');
        container.appendChild(sep);
      } else if ('submenu' in entry || 'widget' in entry) {
        const wrap = document.createElement('div');
        wrap.className = 'bb-menu-sub';
        const { item } = makeRow(entry.label, { hasSub: true });
        item.setAttribute('aria-haspopup', 'true');
        const flyout = document.createElement('div');
        flyout.className = 'bb-menu bb-submenu';
        flyout.setAttribute('role', 'menu');
        if ('widget' in entry) {
          flyout.classList.add('bb-submenu-widget');
          // Built on reveal, and rebuilt on every reveal. A widget that mirrors
          // document state (the page-setup pickers preview the live geometry
          // and check the active preset) would otherwise be a snapshot taken
          // at mount — stale after the first edit, and evaluated before a
          // document has even loaded.
          const build = () => flyout.replaceChildren(entry.widget(() => close()));
          wrap.addEventListener('mouseenter', build);
          wrap.addEventListener('focusin', build);
        } else {
          buildEntries(entry.submenu, flyout);
        }
        wrap.append(item, flyout);
        container.appendChild(wrap);
      } else if ('command' in entry) {
        const cmd = editor.commands.get(entry.command);
        if (!cmd) continue; // skip commands the schema/registry doesn't provide
        const { item, check } = makeRow(entry.label ?? labels[entry.command] ?? entry.command, {});
        item.setAttribute('role', 'menuitemcheckbox');
        item.addEventListener('click', () => {
          if (latest) editor.commands.get(entry.command)?.run(latest, (tr) => editor.dispatch(tr));
          close();
          editor.focus();
        });
        checks.push({ el: check, active: (s) => cmd.isActive?.(s) ?? false });
        if (cmd.isEnabled) enables.push({ el: item, enabled: (s) => cmd.isEnabled!(s) });
        container.appendChild(item);
      } else {
        // ActionEntry
        const { item, check } = makeRow(entry.label, { shortcut: entry.shortcut });
        item.setAttribute('role', entry.isActive ? 'menuitemcheckbox' : 'menuitem');
        item.addEventListener('click', () => {
          entry.run();
          close();
        });
        if (entry.isActive) checks.push({ el: check, active: () => entry.isActive!() });
        if (entry.isEnabled) enables.push({ el: item, enabled: () => entry.isEnabled!() });
        container.appendChild(item);
      }
    }
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

    title.addEventListener('mousedown', (e) => e.preventDefault());
    title.addEventListener('click', () => (openIdx === idx ? close() : open(idx)));
    title.addEventListener('mouseenter', () => {
      if (openIdx >= 0 && openIdx !== idx) open(idx);
    });
    title.addEventListener('keydown', (e) => {
      // Vertical titles open to the right, so ArrowRight also opens there.
      if (e.key === 'ArrowDown' || e.key === 'Enter' || e.key === ' ' || (vertical && e.key === 'ArrowRight')) {
        e.preventDefault();
        open(idx, true);
      }
    });

    buildEntries(menu.entries, dropdown);
    wrap.append(title, dropdown);
    root.appendChild(wrap);
    panels.push({ title, dropdown });
  });

  host.appendChild(root);

  const onDocPointer = (e: Event): void => {
    if (openIdx >= 0 && !root.contains(e.target as Node)) close();
  };
  const onKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape' && openIdx >= 0) {
      e.preventDefault();
      close(true);
    }
  };
  document.addEventListener('pointerdown', onDocPointer);
  document.addEventListener('keydown', onKey);

  function refresh(state: EditorStateOf): void {
    latest = state;
    if (openIdx < 0) return; // only matters while a menu is visible
    for (const c of checks) c.el.textContent = c.active(state) ? '✓' : '';
    for (const e of enables) e.el.disabled = !e.enabled(state);
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
      document.removeEventListener('keydown', onKey);
      root.remove();
    },
  };
}
