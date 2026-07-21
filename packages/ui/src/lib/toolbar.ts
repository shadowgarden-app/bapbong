import type { Collection, Command } from '@shadow-garden/bapbong-contracts';
import {
  type EditorHandle,
  type EditorStateOf,
  injectStyle,
} from './internal.js';

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

/** A colour-picker control: a button (a glyph over a colour bar) that opens a
 *  swatch palette. The host owns the command; the lib renders the popover. */
export interface ToolbarColor {
  kind: 'color';
  title: string;
  /** Glyph above the colour bar (e.g. 'A' for text colour, '🖉' for highlight). */
  glyph: string;
  /** Swatch colours offered (hex). */
  swatches: string[];
  /** Offer a "None / automatic" entry that clears the colour. */
  allowNone?: boolean;
  /** Current colour from state (shown in the bar), or null. */
  value: (state: EditorStateOf) => string | null;
  /** Called with the picked colour (null = cleared). */
  onSelect: (color: string | null) => void;
}

/** A toolbar group entry: a command name (button) or a control (select/colour). */
export type ToolbarEntry = string | ToolbarSelect | ToolbarColor;

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
  spans
    .map(
      ([x1, x2], i) =>
        `<line x1="${x1}" y1="${4 + i * 4}" x2="${x2}" y2="${4 + i * 4}"/>`,
    )
    .join('') +
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
  'clear-format': {
    title: 'Clear formatting',
    svg: '<svg viewBox="0 0 16 16" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12 8 4l3 8M6 9.4h4"/><path d="M2.5 14 13.5 3" opacity=".55"/></svg>',
  },
  bold: { title: 'Bold', label: 'B', className: 'bb-i-bold' },
  italic: { title: 'Italic', label: 'I', className: 'bb-i-italic' },
  underline: { title: 'Underline', label: 'U', className: 'bb-i-underline' },
  strike: { title: 'Strikethrough', label: 'S', className: 'bb-i-strike' },
  superscript: {
    title: 'Superscript',
    svg: '<svg viewBox="0 0 16 16" width="16" height="16" fill="currentColor" aria-hidden="true"><text x="0" y="13" font-size="10.5" font-family="serif">x</text><text x="8.5" y="6.5" font-size="7" font-family="serif">2</text></svg>',
  },
  subscript: {
    title: 'Subscript',
    svg: '<svg viewBox="0 0 16 16" width="16" height="16" fill="currentColor" aria-hidden="true"><text x="0" y="11.5" font-size="10.5" font-family="serif">x</text><text x="8.5" y="15.5" font-size="7" font-family="serif">2</text></svg>',
  },
  'align-left': {
    title: 'Align left',
    svg: alignSvg([
      [2, 14],
      [2, 9],
      [2, 12],
    ]),
  },
  'align-center': {
    title: 'Center',
    svg: alignSvg([
      [2, 14],
      [4, 12],
      [3, 13],
    ]),
  },
  'align-right': {
    title: 'Align right',
    svg: alignSvg([
      [2, 14],
      [7, 14],
      [4, 14],
    ]),
  },
  'align-justify': {
    title: 'Justify',
    svg: alignSvg([
      [2, 14],
      [2, 14],
      [2, 14],
    ]),
  },
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
.bb-toolbar-wrap{position:relative}
.bb-toolbar{display:flex;gap:10px;align-items:center;flex-wrap:nowrap;overflow:hidden;padding:6px 8px;font-family:var(--bb-ui-font,system-ui,-apple-system,sans-serif);color:var(--bb-ui-fg,#2c2c2a);background:var(--bb-ui-bg,#fff);border-bottom:1px solid var(--bb-ui-border,#e3e3e0);box-sizing:border-box}
.bb-toolbar *{box-sizing:border-box}
.bb-toolbar-group{display:flex;gap:2px;flex:none}
.bb-toolbar-group+.bb-toolbar-group{padding-left:10px;border-left:1px solid var(--bb-ui-border,#e3e3e0)}
.bb-toolbar-more{margin-left:auto;flex:none}
.bb-toolbar-pop{position:absolute;z-index:1200;top:100%;left:0;right:0;margin-top:4px;display:flex;flex-wrap:wrap;gap:10px;row-gap:6px;align-items:center;padding:6px 8px;background:var(--bb-ui-menu-bg,#fff);-webkit-backdrop-filter:var(--bb-ui-pop-filter,none);backdrop-filter:var(--bb-ui-pop-filter,none);border:1px solid var(--bb-ui-pop-border,var(--bb-ui-border,#e3e3e0));border-radius:10px;box-shadow:0 8px 28px rgba(0,0,0,.16);font-family:var(--bb-ui-font,system-ui,-apple-system,sans-serif);color:var(--bb-ui-fg,#2c2c2a)}
.bb-toolbar-pop[hidden]{display:none}
.bb-toolbar-btn{min-width:30px;height:30px;display:inline-flex;align-items:center;justify-content:center;border:1px solid transparent;border-radius:6px;background:transparent;color:inherit;cursor:pointer;font-size:14px;line-height:1;padding:0 7px;font-family:inherit}
.bb-toolbar-btn:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-toolbar-btn.is-active{background:var(--bb-ui-active-bg,#e6f1fb);color:var(--bb-ui-active-fg,#0c447c);border-color:var(--bb-ui-active-border,#b5d4f4)}
.bb-toolbar-btn:disabled{opacity:.38;cursor:default}
.bb-toolbar-select{height:30px;border:1px solid var(--bb-ui-border,#e3e3e0);border-radius:6px;background:var(--bb-ui-bg,#fff);color:inherit;font-family:inherit;font-size:13px;padding:0 6px;cursor:pointer}
.bb-toolbar-select:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-toolbar-color{position:relative}
.bb-toolbar-color .bb-color-glyph{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:1px;line-height:1;font-size:13px}
.bb-toolbar-color .bb-color-bar{width:16px;height:3px;border-radius:1px;background:currentColor}
.bb-color-pop{position:absolute;z-index:1200;display:grid;grid-template-columns:repeat(8,16px);gap:4px;padding:8px;background:var(--bb-ui-menu-bg,#fff);-webkit-backdrop-filter:var(--bb-ui-pop-filter,none);backdrop-filter:var(--bb-ui-pop-filter,none);border:1px solid var(--bb-ui-pop-border,var(--bb-ui-border,#e3e3e0));border-radius:8px;box-shadow:0 8px 28px rgba(0,0,0,.16)}
.bb-color-pop[hidden]{display:none}
.bb-color-swatch{width:16px;height:16px;padding:0;border:1px solid rgba(0,0,0,.18);border-radius:3px;cursor:pointer}
.bb-color-none{grid-column:1/-1;font:inherit;font-size:12px;padding:3px 0;border:1px solid var(--bb-ui-border,#e3e3e0);border-radius:5px;background:var(--bb-ui-bg,#fff);cursor:pointer}
.bb-color-none:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-i-bold{font-weight:700}.bb-i-italic{font-style:italic}.bb-i-underline{text-decoration:underline}.bb-i-strike{text-decoration:line-through}
`;

/**
 * Default grouping: every non-`align-*` command, then the `align-*` commands —
 * derived from the registry so new commands appear without config.
 */
export function defaultToolbarGroups(
  commands: Collection<Command>,
): string[][] {
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

  // wrap (position anchor) > toolbar row + overflow popover. The row never
  // wraps: groups that don't fit fold into the popover behind a ⋮ button.
  const wrap = document.createElement('div');
  wrap.className = 'bb-toolbar-wrap';
  const root = document.createElement('div');
  root.className = 'bb-toolbar';
  root.setAttribute('role', 'toolbar');

  const buttons: Array<{ name: string; el: HTMLButtonElement }> = [];
  const selects: Array<{ spec: ToolbarSelect; el: HTMLSelectElement }> = [];
  const colors: Array<{ spec: ToolbarColor; bar: HTMLElement }> = [];
  let latest: EditorStateOf | null = null;

  for (const group of groups) {
    // Keep command entries only if the command exists; controls always render.
    const entries = group.filter(
      (e) => typeof e !== 'string' || editor.commands.has(e),
    );
    if (entries.length === 0) continue;
    const groupEl = document.createElement('div');
    groupEl.className = 'bb-toolbar-group';
    for (const entry of entries) {
      if (typeof entry !== 'string' && entry.kind === 'select') {
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
      if (typeof entry !== 'string' && entry.kind === 'color') {
        const colorWrap = document.createElement('div');
        colorWrap.className = 'bb-toolbar-color';
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'bb-toolbar-btn';
        btn.title = entry.title;
        btn.setAttribute('aria-label', entry.title);
        const glyph = document.createElement('span');
        glyph.className = 'bb-color-glyph';
        glyph.innerHTML = entry.glyph; // text (e.g. 'A') or inline SVG (host-trusted)
        const bar = document.createElement('span');
        bar.className = 'bb-color-bar';
        glyph.appendChild(bar);
        btn.appendChild(glyph);
        const pop = document.createElement('div');
        pop.className = 'bb-color-pop';
        pop.hidden = true;
        const pick = (color: string | null) => {
          entry.onSelect(color);
          pop.hidden = true;
          editor.focus();
        };
        for (const c of entry.swatches) {
          const sw = document.createElement('button');
          sw.type = 'button';
          sw.className = 'bb-color-swatch';
          sw.style.background = c;
          sw.title = c;
          sw.addEventListener('mousedown', (e) => e.preventDefault());
          sw.addEventListener('click', () => pick(c));
          pop.appendChild(sw);
        }
        if (entry.allowNone) {
          const none = document.createElement('button');
          none.type = 'button';
          none.className = 'bb-color-none';
          none.textContent = 'None';
          none.addEventListener('mousedown', (e) => e.preventDefault());
          none.addEventListener('click', () => pick(null));
          pop.appendChild(none);
        }
        btn.addEventListener('mousedown', (e) => e.preventDefault());
        btn.addEventListener('click', () => {
          const open = pop.hidden;
          wrap
            .querySelectorAll('.bb-color-pop')
            .forEach((p) => ((p as HTMLElement).hidden = true));
          pop.hidden = !open;
          if (!pop.hidden) {
            // Position under the button, relative to the outer wrap — the pop
            // lives OUTSIDE the overflow-hidden .bb-toolbar row (which would
            // clip it entirely; the "⋮" popover escapes the same way), and the
            // button itself may sit in the row or in the folded popover.
            const wrapRect = wrap.getBoundingClientRect();
            const btnRect = btn.getBoundingClientRect();
            pop.style.top = `${btnRect.bottom - wrapRect.top + 4}px`;
            const left = Math.max(
              0,
              Math.min(
                btnRect.left - wrapRect.left,
                wrap.clientWidth - pop.offsetWidth - 4,
              ),
            );
            pop.style.left = `${left}px`;
            const onDoc = (ev: Event) => {
              if (
                !colorWrap.contains(ev.target as Node) &&
                !pop.contains(ev.target as Node)
              ) {
                pop.hidden = true;
                document.removeEventListener('pointerdown', onDoc, true);
              }
            };
            document.addEventListener('pointerdown', onDoc, true);
          }
        });
        colorWrap.appendChild(btn);
        wrap.appendChild(pop);
        groupEl.appendChild(colorWrap);
        colors.push({ spec: entry, bar });
        continue;
      }
      const item = items[entry] ?? { title: entry };
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className =
        'bb-toolbar-btn' + (item.className ? ` ${item.className}` : '');
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

  // ── Overflow: single-row toolbar with a ⋮ popover ────────────────────
  // On mount + resize, measure the row; while it overflows, fold whole groups
  // (tail-first) into a popover as wide as the toolbar. Group elements only
  // MOVE between containers, so every button/select/color ref stays live.
  const moreBtn = document.createElement('button');
  moreBtn.type = 'button';
  moreBtn.className = 'bb-toolbar-btn bb-toolbar-more';
  moreBtn.title = 'More tools';
  moreBtn.setAttribute('aria-label', 'More tools');
  moreBtn.setAttribute('aria-expanded', 'false');
  moreBtn.innerHTML =
    '<svg viewBox="0 0 16 16" width="15" height="15" fill="currentColor" aria-hidden="true"><circle cx="8" cy="3.2" r="1.4"/><circle cx="8" cy="8" r="1.4"/><circle cx="8" cy="12.8" r="1.4"/></svg>';
  moreBtn.style.display = 'none';
  root.appendChild(moreBtn);

  const pop = document.createElement('div');
  pop.className = 'bb-toolbar-pop';
  pop.hidden = true;
  wrap.append(root, pop);
  host.appendChild(wrap);

  const closePop = (): void => {
    pop.hidden = true;
    moreBtn.setAttribute('aria-expanded', 'false');
  };
  const onDocPointer = (e: Event): void => {
    if (!pop.hidden && !wrap.contains(e.target as Node)) closePop();
  };
  document.addEventListener('pointerdown', onDocPointer, true);
  moreBtn.addEventListener('mousedown', (e) => e.preventDefault());
  moreBtn.addEventListener('click', () => {
    pop.hidden = !pop.hidden;
    moreBtn.setAttribute('aria-expanded', String(!pop.hidden));
  });

  const fits = (): boolean => root.scrollWidth <= root.clientWidth + 1;
  const layout = (): void => {
    // Unfold everything (in order), then refold the tail until the row fits.
    while (pop.firstChild) root.insertBefore(pop.firstChild, moreBtn);
    moreBtn.style.display = 'none';
    closePop();
    // Refolding moves the anchor buttons — an open color pop would float
    // detached from its button, so close them all.
    wrap
      .querySelectorAll('.bb-color-pop')
      .forEach((p) => ((p as HTMLElement).hidden = true));
    if (fits()) return;
    moreBtn.style.display = '';
    const groupEls = Array.from(root.children).filter((el) =>
      el.classList.contains('bb-toolbar-group'),
    );
    for (let i = groupEls.length - 1; i >= 0 && !fits(); i--) {
      pop.insertBefore(groupEls[i], pop.firstChild);
    }
  };
  // Direct call (no rAF): backgrounded/headless pages throttle rAF, and our
  // mutations never change the row's own box, so there's no observer loop.
  let ro: ResizeObserver | null = null;
  if (typeof ResizeObserver !== 'undefined') {
    ro = new ResizeObserver(layout);
    ro.observe(root);
  }
  window.addEventListener('resize', layout); // fallback where RO misbehaves
  layout();

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
    // Empty → CSS `currentColor` (the button's text colour) shows in the bar.
    for (const { spec, bar } of colors)
      bar.style.background = spec.value(state) ?? '';
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
      ro?.disconnect();
      window.removeEventListener('resize', layout);
      document.removeEventListener('pointerdown', onDocPointer, true);
      wrap.remove();
    },
  };
}
