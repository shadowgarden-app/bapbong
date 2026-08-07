import { colorButton } from './color-picker.js';
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

/** A colour control: a glyph over a bar of the current colour, opening the
 *  shared picker. The host owns the command; the palette belongs to the picker,
 *  not to the caller — every shell offering its own list is how the two shells
 *  ended up with different colours in the same menu. */
export interface ToolbarColor {
  kind: 'color';
  title: string;
  /** Glyph above the colour bar (e.g. 'A' for text colour, '🖉' for highlight). */
  glyph: string;
  /** Wording for the "no colour" row ("Automatic", "No highlight"); omit to
   *  drop that row entirely. */
  clearLabel?: string;
  /** Current colour from state (shown in the bar), or null. */
  value: (state: EditorStateOf) => string | null;
  /** Called with the picked colour (null = cleared). */
  onSelect: (color: string | null) => void;
}

/** One style card in a split button's dropdown: three preview markers, one per
 *  nesting level (e.g. ['1.', 'a.', 'i.']), rendered beside placeholder bars. */
export interface ToolbarSplitOption {
  value: string;
  rows: [string, string, string];
  title?: string;
}

/** A split button: the main button runs the command `name` (toggle, with its
 *  active/enabled tracking and icon), and a narrow ▾ opens a dropdown of style
 *  cards. The host owns the parameterized command each card runs. */
export interface ToolbarSplit {
  kind: 'split';
  /** Command for the main button — also the icon/tooltip lookup key. */
  name: string;
  options: ToolbarSplitOption[];
  /** Currently applied option value (highlights its card), or null. */
  value: (state: EditorStateOf) => string | null;
  /** Called when the user picks a card. */
  onSelect: (value: string) => void;
}

/** A toolbar group entry: a command name (button) or a control (select/colour/split). */
export type ToolbarEntry = string | ToolbarSelect | ToolbarColor | ToolbarSplit;

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
.bb-i-bold{font-weight:700}.bb-i-italic{font-style:italic}.bb-i-underline{text-decoration:underline}.bb-i-strike{text-decoration:line-through}
.bb-toolbar-split{display:flex;align-items:center;gap:0}
.bb-toolbar-split .bb-toolbar-btn{min-width:26px;padding:0 5px}
.bb-split-arrow{min-width:14px !important;width:15px;padding:0 !important}
.bb-split-arrow svg{display:block}
.bb-split-pop{position:absolute;z-index:1200;display:grid;grid-template-columns:repeat(3,auto);gap:8px;padding:10px;background:var(--bb-ui-menu-bg,#fff);-webkit-backdrop-filter:var(--bb-ui-pop-filter,none);backdrop-filter:var(--bb-ui-pop-filter,none);border:1px solid var(--bb-ui-pop-border,var(--bb-ui-border,#e3e3e0));border-radius:10px;box-shadow:0 8px 28px rgba(0,0,0,.16)}
.bb-split-pop[hidden]{display:none}
.bb-split-card{display:flex;flex-direction:column;gap:6px;width:76px;padding:9px 8px;border:1px solid var(--bb-ui-border,#e3e3e0);border-radius:6px;background:transparent;cursor:pointer;font-family:inherit}
.bb-split-card:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-split-card.is-selected{background:var(--bb-ui-active-bg,#e6f1fb);border-color:var(--bb-ui-active-border,#b5d4f4)}
.bb-split-row{display:flex;align-items:center;gap:4px}
.bb-split-row[data-lvl="1"]{padding-left:9px}
.bb-split-row[data-lvl="2"]{padding-left:18px}
.bb-split-marker{flex:none;min-width:14px;font-size:9px;line-height:1.1;text-align:right;color:var(--bb-ui-fg,#2c2c2a);white-space:nowrap}
.bb-split-bar{flex:1;height:3px;border-radius:2px;background:var(--bb-ui-border,#d9d9d4)}
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
  const colors: Array<{ sync(): void }> = [];
  const splits: Array<{ spec: ToolbarSplit; cards: HTMLButtonElement[] }> = [];
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
        // The picker is a body-level fixed panel, so unlike the popover this
        // replaces it needs no wrap-relative arithmetic to escape the row's
        // `overflow: hidden`.
        const ctl = colorButton({
          title: entry.title,
          glyph: entry.glyph,
          clearLabel: entry.clearLabel,
          // `latest`, not `editor.state`: the toolbar mounts before any
          // document exists, and reading state then throws.
          value: () => (latest ? entry.value(latest) : null),
          onPick: (color) => {
            entry.onSelect(color);
            editor.focus();
          },
        });
        groupEl.appendChild(ctl.el);
        colors.push(ctl);
        continue;
      }
      if (typeof entry !== 'string' && entry.kind === 'split') {
        const splitWrap = document.createElement('div');
        splitWrap.className = 'bb-toolbar-split';
        const item = items[entry.name] ?? { title: entry.name };
        // Main button: behaves exactly like a plain command button.
        const main = document.createElement('button');
        main.type = 'button';
        main.className = 'bb-toolbar-btn';
        main.title = item.title;
        main.setAttribute('aria-label', item.title);
        if (item.svg) main.innerHTML = item.svg;
        else main.textContent = item.label ?? entry.name;
        main.addEventListener('mousedown', (e) => e.preventDefault());
        main.addEventListener('click', () => {
          if (!latest) return;
          editor.commands
            .get(entry.name)
            ?.run(latest, (tr) => editor.dispatch(tr));
          editor.focus();
        });
        buttons.push({ name: entry.name, el: main });
        // ▾ arrow: opens the style-card dropdown (portaled to the wrap, like
        // the colour pop, so the overflow-hidden row can't clip it).
        const arrow = document.createElement('button');
        arrow.type = 'button';
        arrow.className = 'bb-toolbar-btn bb-split-arrow';
        arrow.title = `${item.title} styles`;
        arrow.setAttribute('aria-label', `${item.title} styles`);
        arrow.setAttribute('aria-haspopup', 'true');
        arrow.innerHTML =
          '<svg viewBox="0 0 8 8" width="8" height="8" fill="currentColor" aria-hidden="true"><path d="M1 2.5h6L4 6z"/></svg>';
        const pop = document.createElement('div');
        pop.className = 'bb-split-pop';
        pop.hidden = true;
        const cards: HTMLButtonElement[] = [];
        for (const opt of entry.options) {
          const card = document.createElement('button');
          card.type = 'button';
          card.className = 'bb-split-card';
          card.dataset['value'] = opt.value;
          if (opt.title) card.title = opt.title;
          opt.rows.forEach((marker, i) => {
            const row = document.createElement('span');
            row.className = 'bb-split-row';
            row.dataset['lvl'] = String(i);
            const m = document.createElement('span');
            m.className = 'bb-split-marker';
            m.textContent = marker;
            const bar = document.createElement('span');
            bar.className = 'bb-split-bar';
            row.append(m, bar);
            card.appendChild(row);
          });
          card.addEventListener('mousedown', (e) => e.preventDefault());
          card.addEventListener('click', () => {
            entry.onSelect(opt.value);
            pop.hidden = true;
            editor.focus();
          });
          pop.appendChild(card);
          cards.push(card);
        }
        arrow.addEventListener('mousedown', (e) => e.preventDefault());
        arrow.addEventListener('click', () => {
          const open = pop.hidden;
          wrap
            .querySelectorAll('.bb-split-pop')
            .forEach((p) => ((p as HTMLElement).hidden = true));
          pop.hidden = !open;
          if (!pop.hidden) {
            const wrapRect = wrap.getBoundingClientRect();
            const btnRect = arrow.getBoundingClientRect();
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
                !splitWrap.contains(ev.target as Node) &&
                !pop.contains(ev.target as Node)
              ) {
                pop.hidden = true;
                document.removeEventListener('pointerdown', onDoc, true);
              }
            };
            document.addEventListener('pointerdown', onDoc, true);
          }
        });
        splitWrap.append(main, arrow);
        wrap.appendChild(pop);
        groupEl.appendChild(splitWrap);
        splits.push({ spec: entry, cards });
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
    // Refolding moves the anchor buttons — an open color/split pop would float
    // detached from its button, so close them all.
    wrap
      .querySelectorAll('.bb-split-pop')
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
    for (const c of colors) c.sync();
    for (const { spec, cards } of splits) {
      const selected = spec.value(state);
      for (const card of cards)
        card.classList.toggle(
          'is-selected',
          selected != null && card.dataset['value'] === selected,
        );
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
      ro?.disconnect();
      window.removeEventListener('resize', layout);
      document.removeEventListener('pointerdown', onDocPointer, true);
      wrap.remove();
    },
  };
}
