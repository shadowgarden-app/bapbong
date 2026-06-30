import type { Collection, Command } from '@shadow-garden/bapbong-contracts';
import { type EditorHandle, type EditorStateOf, injectStyle } from './internal.js';

/** Presentation for one toolbar button — the headless {@link Command} carries
 *  no UI, so labels/icons live here. */
export interface ToolbarItem {
  /** Tooltip + accessible label. */
  title: string;
  /** Text/glyph shown when `svg` is absent. */
  label?: string;
  /** Inline SVG markup (used instead of `label`). */
  svg?: string;
  /** Extra class on the button (e.g. to render a bold "B"). */
  className?: string;
}

/** A dropdown control in the toolbar (e.g. font size / family). The host owns
 *  the command to run — the toolbar just renders the `<select>` and reports the
 *  current value, so the lib stays decoupled from specific commands. */
export interface ToolbarSelect {
  kind: 'select';
  /** Tooltip + accessible label. */
  title: string;
  options: { label: string; value: string }[];
  /** Current value to show from editor state (`''` = none / mixed). */
  value: (state: EditorStateOf) => string;
  /** Called when the user picks a value. */
  onSelect: (value: string) => void;
  /** Fixed width in px (defaults to auto). */
  width?: number;
}

/** A toolbar group entry: a command name (button) or a control like a select. */
export type ToolbarEntry = string | ToolbarSelect;

export interface ToolbarOptions {
  /** Groups rendered as separated clusters — command names (buttons) and/or
   *  controls (selects). Defaults to marks then alignments, from the registry. */
  groups?: ToolbarEntry[][];
  /** Presentation per command name, merged over the built-in defaults. */
  items?: Record<string, ToolbarItem>;
}

/** Handle returned by {@link mountToolbar}. */
export interface ToolbarHandle {
  destroy(): void;
}

/** Three horizontal rules with the given x-spans — a tiny alignment glyph. */
const alignSvg = (spans: Array<[number, number]>) =>
  `<svg width="15" height="15" viewBox="0 0 16 16" aria-hidden="true">` +
  `<g stroke="currentColor" stroke-width="1.6" stroke-linecap="round">` +
  spans.map(([x1, x2], i) => `<line x1="${x1}" y1="${4 + i * 4}" x2="${x2}" y2="${4 + i * 4}"/>`).join('') +
  `</g></svg>`;

const DEFAULT_ITEMS: Record<string, ToolbarItem> = {
  undo: {
    title: 'Undo',
    svg: '<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h6a3.5 3.5 0 0 1 0 7H6"/><path d="M4 7 7 4M4 7l3 3"/></svg>',
  },
  redo: {
    title: 'Redo',
    svg: '<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12 7H6a3.5 3.5 0 0 0 0 7h4"/><path d="m12 7-3-3m3 3-3 3"/></svg>',
  },
  bold: { title: 'Bold', label: 'B', className: 'bb-i-bold' },
  italic: { title: 'Italic', label: 'I', className: 'bb-i-italic' },
  underline: { title: 'Underline', label: 'U', className: 'bb-i-underline' },
  strike: { title: 'Strikethrough', label: 'S', className: 'bb-i-strike' },
  'align-left': { title: 'Align left', svg: alignSvg([[2, 14], [2, 9], [2, 12]]) },
  'align-center': { title: 'Center', svg: alignSvg([[2, 14], [4, 12], [3, 13]]) },
  'align-right': { title: 'Align right', svg: alignSvg([[2, 14], [7, 14], [4, 14]]) },
  'align-justify': { title: 'Justify', svg: alignSvg([[2, 14], [2, 14], [2, 14]]) },
  'bullet-list': {
    title: 'Bullet list',
    svg: '<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"><path d="M6 4h8M6 8h8M6 12h8"/><circle cx="3" cy="4" r="1" fill="currentColor" stroke="none"/><circle cx="3" cy="8" r="1" fill="currentColor" stroke="none"/><circle cx="3" cy="12" r="1" fill="currentColor" stroke="none"/></svg>',
  },
  'ordered-list': {
    title: 'Numbered list',
    svg: '<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"><path d="M6 4h8M6 8h8M6 12h8"/><text x="0.5" y="5.6" font-size="5.5" fill="currentColor" stroke="none">1</text><text x="0.5" y="9.6" font-size="5.5" fill="currentColor" stroke="none">2</text><text x="0.5" y="13.6" font-size="5.5" fill="currentColor" stroke="none">3</text></svg>',
  },
};

const STYLE = `
.bb-toolbar{display:flex;gap:10px;align-items:center;flex-wrap:wrap;padding:6px 8px;font-family:var(--bb-ui-font,system-ui,-apple-system,sans-serif);color:var(--bb-ui-fg,#2c2c2a);background:var(--bb-ui-bg,#fff);border-bottom:1px solid var(--bb-ui-border,#e3e3e0);box-sizing:border-box}
.bb-toolbar *{box-sizing:border-box}
.bb-toolbar-group{display:flex;gap:2px}
.bb-toolbar-group+.bb-toolbar-group{padding-left:10px;border-left:1px solid var(--bb-ui-border,#e3e3e0)}
.bb-toolbar-btn{min-width:30px;height:30px;display:inline-flex;align-items:center;justify-content:center;border:1px solid transparent;border-radius:6px;background:transparent;color:inherit;cursor:pointer;font-size:14px;line-height:1;padding:0 7px;font-family:inherit}
.bb-toolbar-btn:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-toolbar-btn.is-active{background:var(--bb-ui-active-bg,#e6f1fb);color:var(--bb-ui-active-fg,#0c447c);border-color:var(--bb-ui-active-border,#b5d4f4)}
.bb-toolbar-btn:disabled{opacity:.38;cursor:default}
.bb-toolbar-select{height:30px;border:1px solid var(--bb-ui-border,#e3e3e0);border-radius:6px;background:var(--bb-ui-bg,#fff);color:inherit;font-family:inherit;font-size:13px;padding:0 6px;cursor:pointer}
.bb-toolbar-select:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-i-bold{font-weight:700}.bb-i-italic{font-style:italic}.bb-i-underline{text-decoration:underline}.bb-i-strike{text-decoration:line-through}
`;

/**
 * Default grouping: every non-`align-*` command, then the `align-*` commands —
 * derived from the registry so new commands appear without config.
 */
export function defaultToolbarGroups(commands: Collection<Command>): string[][] {
  const names = [...commands].map((c) => c.name);
  const aligns = names.filter((n) => n.startsWith('align-'));
  const rest = names.filter((n) => !n.startsWith('align-'));
  return [rest, aligns].filter((g) => g.length > 0);
}

/**
 * Render a toolbar into `host` and wire it to `editor`. The lib owns all the
 * DOM, styling and active-state tracking; a host framework only provides the
 * element and the editor handle, then calls `destroy()` on teardown.
 */
export function mountToolbar(
  host: HTMLElement,
  editor: EditorHandle,
  options: ToolbarOptions = {},
): ToolbarHandle {
  injectStyle('bb-ui-toolbar-styles', STYLE);
  const items = { ...DEFAULT_ITEMS, ...(options.items ?? {}) };
  const groups = options.groups ?? defaultToolbarGroups(editor.commands);

  const root = document.createElement('div');
  root.className = 'bb-toolbar';
  root.setAttribute('role', 'toolbar');

  const buttons: Array<{ name: string; el: HTMLButtonElement }> = [];
  const selects: Array<{ spec: ToolbarSelect; el: HTMLSelectElement }> = [];
  let latest: EditorStateOf | null = null;

  for (const group of groups) {
    // Keep command entries only if the command exists; controls always render.
    const entries = group.filter((e) => typeof e !== 'string' || editor.commands.has(e));
    if (entries.length === 0) continue;
    const groupEl = document.createElement('div');
    groupEl.className = 'bb-toolbar-group';
    for (const entry of entries) {
      if (typeof entry !== 'string') {
        // A control (currently: select dropdown).
        const sel = document.createElement('select');
        sel.className = 'bb-toolbar-select';
        sel.title = entry.title;
        sel.setAttribute('aria-label', entry.title);
        if (entry.width) sel.style.width = `${entry.width}px`;
        for (const opt of entry.options) {
          const o = document.createElement('option');
          o.value = opt.value;
          o.textContent = opt.label;
          sel.appendChild(o);
        }
        sel.addEventListener('mousedown', (e) => e.stopPropagation());
        sel.addEventListener('change', () => {
          entry.onSelect(sel.value);
          editor.focus();
        });
        groupEl.appendChild(sel);
        selects.push({ spec: entry, el: sel });
        continue;
      }
      const item = items[entry] ?? { title: entry };
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'bb-toolbar-btn' + (item.className ? ` ${item.className}` : '');
      btn.title = item.title;
      btn.setAttribute('aria-label', item.title);
      if (item.svg) btn.innerHTML = item.svg;
      else btn.textContent = item.label ?? entry;
      // Keep the editor's selection/focus when a button is pressed.
      btn.addEventListener('mousedown', (e) => e.preventDefault());
      btn.addEventListener('click', () => {
        if (!latest) return;
        editor.commands.get(entry)?.run(latest, (tr) => editor.dispatch(tr));
        editor.focus();
      });
      groupEl.appendChild(btn);
      buttons.push({ name: entry, el: btn });
    }
    root.appendChild(groupEl);
  }
  host.appendChild(root);

  const refresh = (state: EditorStateOf): void => {
    latest = state;
    for (const { name, el } of buttons) {
      const cmd = editor.commands.get(name);
      if (!cmd) continue;
      const active = cmd.isActive?.(state) ?? false;
      el.classList.toggle('is-active', active);
      el.setAttribute('aria-pressed', String(active));
      // No isEnabled → always enabled (a run-probe would wrongly disable
      // idempotent ops like "align-center" when already centered).
      el.disabled = cmd.isEnabled ? !cmd.isEnabled(state) : false;
    }
    for (const { spec, el } of selects) el.value = spec.value(state);
  };

  const off = editor.onChange((c) => refresh(c.state));
  try {
    refresh(editor.state); // reflect a document that's already loaded
  } catch {
    /* no document yet — buttons activate on the first change */
  }

  return {
    destroy() {
      off();
      root.remove();
    },
  };
}
