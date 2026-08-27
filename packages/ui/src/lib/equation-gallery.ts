import type {
  EqNode,
  VectorImageSpec,
  VectorOp,
} from '@shadow-garden/bapbong-contracts';
import { astToLinear } from '@shadow-garden/bapbong-contracts';
import { injectStyle } from './internal.js';

/** One entry of the built-in gallery: a name and the equation it inserts. */
export interface BuiltInEquation {
  name: string;
  ast: EqNode[];
}

/** Characters, letterformed the way every equation spells them (math italic
 *  for variables) — the gallery stores exactly what gets inserted. */
const c = (s: string): EqNode[] => [...s].map((ch) => ({ t: 'chr', ch }));
const sup = (base: EqNode[], exp: string): EqNode => ({
  t: 'scr',
  base,
  sub: [],
  sup: c(exp),
});
const subsup = (base: EqNode[], lo: string, hi: string): EqNode => ({
  t: 'scr',
  base,
  sub: c(lo),
  sup: c(hi),
});

/**
 * The built-in equations Insert ▸ Equation offers, the way Word's gallery
 * does. Chosen for the documents bapbong is actually used on — Vietnamese
 * exam papers — so the list is the secondary-school canon, not Word's
 * (Fourier series has no place in a year-12 paper). Each is a real AST: what
 * lands in the document is editable, exports as m:oMath, and needs no
 * conversion.
 */
export const BUILT_IN_EQUATIONS: readonly BuiltInEquation[] = [
  {
    name: 'Area of a circle',
    ast: [...c('𝐴 = 𝜋'), sup(c('𝑟'), '2')],
  },
  {
    name: 'Pythagorean theorem',
    ast: [
      sup(c('𝑎'), '2'),
      ...c(' + '),
      sup(c('𝑏'), '2'),
      ...c(' = '),
      sup(c('𝑐'), '2'),
    ],
  },
  {
    name: 'Quadratic formula',
    ast: [
      ...c('𝑥 = '),
      {
        t: 'frac',
        num: [
          ...c('−𝑏 ± '),
          {
            t: 'rad',
            deg: [],
            body: [sup(c('𝑏'), '2'), ...c(' − 4𝑎𝑐')],
          },
        ],
        den: c('2𝑎'),
      },
    ],
  },
  {
    name: 'Law of cosines',
    ast: [
      sup(c('𝑎'), '2'),
      ...c(' = '),
      sup(c('𝑏'), '2'),
      ...c(' + '),
      sup(c('𝑐'), '2'),
      ...c(' − 2𝑏𝑐 cos 𝐴'),
    ],
  },
  {
    name: 'Binomial theorem',
    ast: [
      // The exponent belongs ON the bracket: a script with an empty base
      // would render its own empty slot — the dotted placeholder box.
      {
        t: 'scr',
        base: [{ t: 'fence', l: '(', r: ')', body: c('𝑎 + 𝑏') }],
        sub: [],
        sup: c('𝑛'),
      },
      ...c(' = '),
      {
        t: 'big',
        op: '∑',
        lo: c('𝑘 = 0'),
        hi: c('𝑛'),
        body: [
          subsup(c('𝐶'), '𝑛', '𝑘'),
          sup(c('𝑎'), '𝑘'),
          { t: 'scr', base: c('𝑏'), sub: [], sup: c('𝑛−𝑘') },
        ],
      },
    ],
  },
];

const SVG_NS = 'http://www.w3.org/2000/svg';

/**
 * A typeset equation drawn as inline SVG — the SAME display list the canvas
 * painter replays, so a preview is the drawing the page will show rather
 * than an approximation of it. Pure geometry from contracts: this package
 * still knows nothing about the layout engine (the caller typesets and
 * hands the spec over).
 */
export function equationPreviewSvg(
  spec: VectorImageSpec,
  opts: { maxWidth?: number } = {},
): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', `0 0 ${spec.width} ${spec.height}`);
  // Scale to fit the menu, never up: a short equation keeps its real size,
  // a long one shrinks the way it would in a narrower column.
  const max = opts.maxWidth ?? spec.width;
  const scale = Math.min(1, max / Math.max(1, spec.width));
  svg.setAttribute('width', String(Math.round(spec.width * scale)));
  svg.setAttribute('height', String(Math.round(spec.height * scale)));
  svg.setAttribute('aria-hidden', 'true');
  for (const op of spec.ops) svg.append(opElement(op));
  return svg;
}

function opElement(op: VectorOp): SVGElement {
  if (op.kind === 'line') {
    const el = document.createElementNS(SVG_NS, 'line');
    el.setAttribute('x1', String(op.x1));
    el.setAttribute('y1', String(op.y1));
    el.setAttribute('x2', String(op.x2));
    el.setAttribute('y2', String(op.y2));
    el.setAttribute('stroke', op.color);
    el.setAttribute('stroke-width', String(Math.max(op.width, 0.75)));
    return el;
  }
  if (op.kind === 'polygon') {
    const el = document.createElementNS(SVG_NS, 'polygon');
    el.setAttribute('points', op.points.map((p) => `${p.x},${p.y}`).join(' '));
    el.setAttribute('fill', op.fill ?? 'none');
    if (op.stroke) {
      el.setAttribute('stroke', op.stroke);
      el.setAttribute('stroke-width', String(op.strokeWidth ?? 1));
    }
    return el;
  }
  const el = document.createElementNS(SVG_NS, 'text');
  el.setAttribute('x', String(op.x));
  el.setAttribute('y', String(op.y));
  el.setAttribute('font-size', String(op.size));
  el.setAttribute('font-family', op.family);
  el.setAttribute('fill', op.color);
  if (op.bold) el.setAttribute('font-weight', 'bold');
  if (op.italic) el.setAttribute('font-style', 'italic');
  if (op.vAlign === 'top') el.setAttribute('dominant-baseline', 'hanging');
  // Per-character advances: the source kerned by hand, so place each glyph
  // rather than letting the browser measure the run.
  if (op.dx) {
    let x = op.x;
    for (const [i, ch] of [...op.text].entries()) {
      const t = document.createElementNS(SVG_NS, 'tspan');
      t.setAttribute('x', String(x));
      t.textContent = ch;
      el.append(t);
      x += op.dx[i] ?? 0;
    }
  } else {
    el.textContent = op.text;
  }
  return el;
}

export interface EquationGalleryOptions {
  /** Entries to offer. Defaults to {@link BUILT_IN_EQUATIONS}. */
  items?: readonly BuiltInEquation[];
  /**
   * Typesets one entry the way the document does — hosts pass the layout
   * engine bound to the editor's own measurer, so a preview matches the
   * page exactly. Without it the preview falls back to the equation's
   * linear spelling.
   */
  layout?: (ast: EqNode[], sizePt: number) => VectorImageSpec | null;
  /** A gallery entry was chosen — insert this equation. */
  onPick: (ast: EqNode[]) => void;
  /** The footer action: start an empty equation instead. */
  onNew: () => void;
  /** Shortcut hint shown beside the footer action (e.g. "⌥="). */
  newShortcut?: string;
}

const STYLE = `
.bb-eqg{display:flex;flex-direction:column;min-width:236px;max-width:300px;user-select:none}
.bb-eqg-head{font-size:11px;font-weight:600;opacity:.55;padding:2px 8px 6px}
.bb-eqg-list{display:flex;flex-direction:column;max-height:320px;overflow-y:auto}
.bb-eqg-item{display:flex;flex-direction:column;gap:3px;align-items:flex-start;width:100%;padding:7px 9px;border:0;border-radius:6px;background:transparent;color:inherit;font:inherit;text-align:left;cursor:pointer}
.bb-eqg-item:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-eqg-name{font-size:11px;opacity:.6}
/* The preview is the equation TYPESET — the same display list the canvas
   paints. The text form is the fallback for a host that cannot lay out. */
.bb-eqg-prev{display:flex;align-items:center;min-height:22px;max-width:100%;overflow:hidden}
.bb-eqg-prev svg{display:block}
.bb-eqg-prev-text{font-family:"Times New Roman",Tinos,serif;font-size:15px;line-height:1.35;white-space:nowrap;text-overflow:ellipsis}
.bb-eqg-foot{margin-top:4px;padding-top:5px;border-top:1px solid var(--bb-ui-border,#e3e3e0)}
.bb-eqg-new{display:flex;align-items:center;gap:8px;width:100%;height:30px;padding:0 9px;border:0;border-radius:6px;background:transparent;color:inherit;font:inherit;font-size:13px;text-align:left;cursor:pointer}
.bb-eqg-new:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-eqg-key{margin-left:auto;opacity:.5;font-size:12px}
`;

/**
 * Insert ▸ Equation's flyout: a gallery of built-in equations over a footer
 * that starts an empty one — Word's shape, because that is where people look
 * for it. Picking an entry inserts a real equation node; the footer runs the
 * host's insert command.
 */
/** Preview type size and the width the menu allows before scaling down. */
const PREVIEW_PT = 11;
const PREVIEW_MAX_W = 264;

export function equationGallery(options: EquationGalleryOptions): HTMLElement {
  injectStyle('bb-ui-eqg-styles', STYLE);
  const root = document.createElement('div');
  root.className = 'bb-eqg';

  const head = document.createElement('div');
  head.className = 'bb-eqg-head';
  head.textContent = 'Built-in';
  root.append(head);

  const list = document.createElement('div');
  list.className = 'bb-eqg-list';
  for (const item of options.items ?? BUILT_IN_EQUATIONS) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'bb-eqg-item';
    const prev = document.createElement('div');
    prev.className = 'bb-eqg-prev';
    // Typeset when the host can (identical to the page), the linear
    // spelling when it cannot.
    const spec = options.layout?.(item.ast, PREVIEW_PT) ?? null;
    if (spec && spec.ops.length)
      prev.append(equationPreviewSvg(spec, { maxWidth: PREVIEW_MAX_W }));
    else {
      prev.classList.add('bb-eqg-prev-text');
      prev.textContent = astToLinear(item.ast);
    }
    const name = document.createElement('div');
    name.className = 'bb-eqg-name';
    name.textContent = item.name;
    btn.append(prev, name);
    // Keep focus where it is: the menubar rebuilds a widget flyout on
    // `focusin`, so letting the press focus this button would replace the
    // element between mousedown and mouseup — and no click would ever fire.
    // Same guard the table-size and page-setup pickers use.
    btn.addEventListener('mousedown', (e) => e.preventDefault());
    btn.addEventListener('click', () => options.onPick(item.ast));
    list.append(btn);
  }
  root.append(list);

  const foot = document.createElement('div');
  foot.className = 'bb-eqg-foot';
  const neu = document.createElement('button');
  neu.type = 'button';
  neu.className = 'bb-eqg-new';
  const label = document.createElement('span');
  label.textContent = 'Insert new equation';
  neu.append(label);
  if (options.newShortcut) {
    const key = document.createElement('span');
    key.className = 'bb-eqg-key';
    key.textContent = options.newShortcut;
    neu.append(key);
  }
  neu.addEventListener('mousedown', (e) => e.preventDefault());
  neu.addEventListener('click', () => options.onNew());
  foot.append(neu);
  root.append(foot);

  return root;
}
