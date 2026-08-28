import type {
  EqNode,
  EqStructure,
  VectorImageSpec,
} from '@shadow-garden/bapbong-contracts';
import {
  eqStructureGroups,
  MATH_AUTOCORRECT,
  MATH_SYMBOL_SETS,
} from '@shadow-garden/bapbong-contracts';
import { equationPreviewSvg } from './equation-gallery.js';
import { injectStyle, TABS_CSS } from './internal.js';

export interface EquationPanelOptions {
  /**
   * Typesets one template the way the document would — the host passes the
   * layout engine bound to the editor's measurer, so a preview is the
   * drawing the page will show. Without it the structure grid falls back to
   * plain labels.
   */
  layout?: (ast: EqNode[], sizePt: number) => VectorImageSpec | null;
  /** A symbol was chosen — insert this character at the slot caret. */
  onSymbol: (ch: string) => void;
  /** A structure was chosen — insert this template at the slot caret. */
  onStructure: (s: EqStructure) => void;
}

export interface EquationPanel {
  /** The panel element. The caller positions it and shows/hides it. */
  readonly el: HTMLElement;
  destroy(): void;
}

/** `\name` for a glyph, when Math AutoCorrect has one — the palette doubles
 *  as the place people learn the shorthand. First spelling wins, so `\pi`
 *  shows rather than a later alias. */
const SHORTHAND = new Map<string, string>();
for (const [name, ch] of Object.entries(MATH_AUTOCORRECT))
  if (!SHORTHAND.has(ch)) SHORTHAND.set(ch, `\\${name}`);

const PREVIEW_PT = 10;

const STYLE = `
.bb-eqp{width:352px;display:flex;flex-direction:column;background:var(--bb-ui-menu-bg,#fff);-webkit-backdrop-filter:var(--bb-ui-pop-filter,none);backdrop-filter:var(--bb-ui-pop-filter,none);border:1px solid var(--bb-ui-pop-border,var(--bb-ui-border,#e3e3e0));border-radius:10px;box-shadow:0 8px 28px rgba(0,0,0,.14);user-select:none;overflow:hidden}
.bb-eqp-tabs{display:flex;padding:8px 8px 0}
.bb-eqp-body{padding:8px}
.bb-eqp-set{width:100%;height:30px;margin-bottom:7px;border:0.5px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-radius:6px;background:var(--bb-ui-control-bg,#fff);color:inherit;font:inherit;font-size:13px;padding:0 8px}
.bb-eqp-grid{display:grid;grid-template-columns:repeat(10,1fr);gap:1px;max-height:190px;overflow-y:auto}
.bb-eqp-sym{padding:5px 0;border:0;border-radius:5px;background:transparent;color:inherit;font-family:"Cambria Math","Times New Roman",Tinos,serif;font-size:16px;line-height:1.2;cursor:pointer}
.bb-eqp-sym:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-eqp-hint{display:flex;align-items:center;justify-content:space-between;gap:8px;min-height:22px;margin-top:7px;padding-top:6px;border-top:0.5px solid var(--bb-ui-border,#e3e3e0);font-size:12px;opacity:.6}
.bb-eqp-hint-g{font-family:"Cambria Math","Times New Roman",Tinos,serif;font-size:15px}
.bb-eqp-hint-k{font-family:ui-monospace,SFMono-Regular,Menlo,monospace}
.bb-eqp-scroll{max-height:230px;overflow-y:auto}
.bb-eqp-group{font-size:11px;font-weight:600;opacity:.5;padding:6px 2px 3px}
.bb-eqp-row{display:grid;grid-template-columns:repeat(4,1fr);gap:3px}
.bb-eqp-st{display:flex;flex-direction:column;align-items:center;justify-content:flex-end;gap:3px;min-height:52px;padding:6px 2px;border:0;border-radius:6px;background:transparent;color:inherit;font:inherit;cursor:pointer}
.bb-eqp-st:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-eqp-st-name{font-size:10px;opacity:.6;text-align:center;line-height:1.25}
.bb-eqp-st svg{display:block}
`;

/**
 * The equation palette: Word's Symbols and Structures galleries as ONE
 * floating surface with tabs inside it, rather than a bar that opens a
 * separate flyout. Switching tabs then swaps the contents while the frame
 * stays put — and inserting six symbols in a row costs no re-opening.
 *
 * Knows nothing about the editor: it reports a chosen symbol or template and
 * the host hands that to the slot editor.
 */
export function equationPanel(options: EquationPanelOptions): EquationPanel {
  injectStyle('bb-ui-tabs', TABS_CSS);
  injectStyle('bb-ui-eqp-styles', STYLE);

  const root = document.createElement('div');
  root.className = 'bb-eqp';
  root.setAttribute('role', 'group');
  root.setAttribute('aria-label', 'Equation palette');

  // Every press inside the panel must leave the caret where it is: the slot
  // caret lives in the editor, and taking focus would end the edit before
  // the click ever ran. Same guard the galleries use.
  root.addEventListener('mousedown', (e) => e.preventDefault());

  const strip = document.createElement('div');
  strip.className = 'bb-eqp-tabs';
  const tabs = document.createElement('div');
  tabs.className = 'bb-tabs';
  tabs.setAttribute('role', 'tablist');
  strip.append(tabs);

  const body = document.createElement('div');
  body.className = 'bb-eqp-body';

  let current: 'symbols' | 'structures' = 'symbols';

  const tabEls: Record<string, HTMLButtonElement> = {};
  const mkTab = (id: 'symbols' | 'structures', label: string) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'bb-tab';
    b.textContent = label;
    b.setAttribute('role', 'tab');
    b.addEventListener('click', () => {
      current = id;
      render();
    });
    tabEls[id] = b;
    tabs.append(b);
  };
  mkTab('symbols', 'Symbols');
  mkTab('structures', 'Structures');

  root.append(strip, body);

  function symbolsBody(): void {
    const set = document.createElement('select');
    set.className = 'bb-eqp-set';
    set.setAttribute('aria-label', 'Symbol set');
    for (const [i, s] of MATH_SYMBOL_SETS.entries()) {
      const o = document.createElement('option');
      o.value = String(i);
      o.textContent = s.name;
      set.append(o);
    }
    set.value = String(setIndex);

    const grid = document.createElement('div');
    grid.className = 'bb-eqp-grid';

    const hint = document.createElement('div');
    hint.className = 'bb-eqp-hint';
    const hintG = document.createElement('span');
    hintG.className = 'bb-eqp-hint-g';
    const hintK = document.createElement('span');
    hintK.className = 'bb-eqp-hint-k';
    hint.append(hintG, hintK);

    const fill = () => {
      grid.replaceChildren();
      const chars = MATH_SYMBOL_SETS[setIndex]?.chars ?? '';
      for (const ch of chars) {
        const b = document.createElement('button');
        b.type = 'button';
        b.className = 'bb-eqp-sym';
        b.textContent = ch;
        const key = SHORTHAND.get(ch);
        b.title = key ? `${ch}   ${key}` : ch;
        b.addEventListener('mouseenter', () => {
          hintG.textContent = ch;
          hintK.textContent = key ?? '';
        });
        b.addEventListener('click', () => options.onSymbol(ch));
        grid.append(b);
      }
    };
    set.addEventListener('change', () => {
      setIndex = Number(set.value) || 0;
      fill();
    });
    fill();
    body.replaceChildren(set, grid, hint);
  }

  function structuresBody(): void {
    const scroll = document.createElement('div');
    scroll.className = 'bb-eqp-scroll';
    for (const group of eqStructureGroups()) {
      const h = document.createElement('div');
      h.className = 'bb-eqp-group';
      h.textContent = group.name;
      const row = document.createElement('div');
      row.className = 'bb-eqp-row';
      for (const item of group.items) {
        const b = document.createElement('button');
        b.type = 'button';
        b.className = 'bb-eqp-st';
        b.title = `${group.name} — ${item.name}`;
        // The preview IS the template: the same empty-slot boxes the page
        // will draw, so what you point at is what you get.
        const spec = options.layout?.([item.node], PREVIEW_PT) ?? null;
        if (spec && spec.ops.length)
          b.append(equationPreviewSvg(spec, { maxWidth: 60 }));
        const name = document.createElement('span');
        name.className = 'bb-eqp-st-name';
        name.textContent = item.name;
        b.append(name);
        b.addEventListener('click', () => options.onStructure(item));
        row.append(b);
      }
      scroll.append(h, row);
    }
    body.replaceChildren(scroll);
  }

  let setIndex = 0;

  function render(): void {
    for (const [id, el] of Object.entries(tabEls))
      el.setAttribute('aria-selected', String(id === current));
    if (current === 'symbols') symbolsBody();
    else structuresBody();
  }

  render();

  return {
    el: root,
    destroy() {
      root.remove();
    },
  };
}
