import type { Collection, Command, Dispatch, EditorChange } from '@shadow-garden/bapbong-contracts';

/** The editor `run` state type, derived from the {@link Command} contract so
 *  this package needs no direct ProseMirror dependency. */
type EditorStateOf = Parameters<Command['run']>[0];

/**
 * The minimal editor surface the UI binds to. `BapbongEditor` satisfies it
 * structurally, so this package never imports the editor — both just speak the
 * `contracts` vocabulary (the same decoupling as the plugin contract). A host
 * framework only supplies the element + this handle.
 */
export interface EditorHandle {
  /** Named command registry the toolbar renders + dispatches against. */
  readonly commands: Collection<Command>;
  /** Current document + selection (throws before a document is loaded). */
  readonly state: EditorStateOf;
  /** Apply a transaction. */
  dispatch: Dispatch;
  /** Return focus to the editor (after a button takes a click). */
  focus(): void;
  /** Subscribe to editor cycles; returns an unsubscribe. Drives active state. */
  onChange(cb: (c: EditorChange) => void): () => void;
}

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

export interface ToolbarOptions {
  /** Command-name groups, rendered as separated button clusters. Defaults to
   *  marks then alignments, derived from the registry. */
  groups?: string[][];
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
  bold: { title: 'Bold', label: 'B', className: 'bb-i-bold' },
  italic: { title: 'Italic', label: 'I', className: 'bb-i-italic' },
  underline: { title: 'Underline', label: 'U', className: 'bb-i-underline' },
  strike: { title: 'Strikethrough', label: 'S', className: 'bb-i-strike' },
  'align-left': { title: 'Align left', svg: alignSvg([[2, 14], [2, 9], [2, 12]]) },
  'align-center': { title: 'Center', svg: alignSvg([[2, 14], [4, 12], [3, 13]]) },
  'align-right': { title: 'Align right', svg: alignSvg([[2, 14], [7, 14], [4, 14]]) },
  'align-justify': { title: 'Justify', svg: alignSvg([[2, 14], [2, 14], [2, 14]]) },
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
.bb-i-bold{font-weight:700}.bb-i-italic{font-style:italic}.bb-i-underline{text-decoration:underline}.bb-i-strike{text-decoration:line-through}
`;

let stylesInjected = false;
function injectStyles(): void {
  if (stylesInjected || document.getElementById('bb-ui-toolbar-styles')) {
    stylesInjected = true;
    return;
  }
  const el = document.createElement('style');
  el.id = 'bb-ui-toolbar-styles';
  el.textContent = STYLE;
  document.head.appendChild(el);
  stylesInjected = true;
}

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
  injectStyles();
  const items = { ...DEFAULT_ITEMS, ...(options.items ?? {}) };
  const groups = options.groups ?? defaultToolbarGroups(editor.commands);

  const root = document.createElement('div');
  root.className = 'bb-toolbar';
  root.setAttribute('role', 'toolbar');

  const buttons: Array<{ name: string; el: HTMLButtonElement }> = [];
  let latest: EditorStateOf | null = null;

  for (const group of groups) {
    const names = group.filter((n) => editor.commands.has(n));
    if (names.length === 0) continue;
    const groupEl = document.createElement('div');
    groupEl.className = 'bb-toolbar-group';
    for (const name of names) {
      const item = items[name] ?? { title: name };
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'bb-toolbar-btn' + (item.className ? ` ${item.className}` : '');
      btn.title = item.title;
      btn.setAttribute('aria-label', item.title);
      if (item.svg) btn.innerHTML = item.svg;
      else btn.textContent = item.label ?? name;
      // Keep the editor's selection/focus when a button is pressed.
      btn.addEventListener('mousedown', (e) => e.preventDefault());
      btn.addEventListener('click', () => {
        if (!latest) return;
        editor.commands.get(name)?.run(latest, (tr) => editor.dispatch(tr));
        editor.focus();
      });
      groupEl.appendChild(btn);
      buttons.push({ name, el: btn });
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
