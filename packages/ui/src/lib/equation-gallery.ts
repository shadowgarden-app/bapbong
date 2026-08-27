import type { EqNode } from '@shadow-garden/bapbong-contracts';
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
 * (Fourier series has no place in a lớp 12 paper). Each is a real AST: what
 * lands in the document is editable, exports as m:oMath, and needs no
 * conversion.
 */
export const BUILT_IN_EQUATIONS: readonly BuiltInEquation[] = [
  {
    name: 'Diện tích hình tròn',
    ast: [...c('𝐴 = 𝜋'), sup(c('𝑟'), '2')],
  },
  {
    name: 'Định lý Pythagore',
    ast: [
      sup(c('𝑎'), '2'),
      ...c(' + '),
      sup(c('𝑏'), '2'),
      ...c(' = '),
      sup(c('𝑐'), '2'),
    ],
  },
  {
    name: 'Nghiệm phương trình bậc hai',
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
    name: 'Định lý cosin',
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
    name: 'Nhị thức Newton',
    ast: [
      { t: 'fence', l: '(', r: ')', body: c('𝑎 + 𝑏') },
      { t: 'scr', base: [], sub: [], sup: c('𝑛') },
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

export interface EquationGalleryOptions {
  /** Entries to offer. Defaults to {@link BUILT_IN_EQUATIONS}. */
  items?: readonly BuiltInEquation[];
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
/* The preview is the equation's own plain-text spelling, set in the document
   face — the same characters the insert produces, so nothing is promised
   that the document will not show. */
.bb-eqg-prev{font-family:"Times New Roman",Tinos,serif;font-size:15px;line-height:1.35;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%}
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
export function equationGallery(options: EquationGalleryOptions): HTMLElement {
  injectStyle('bb-ui-eqg-styles', STYLE);
  const root = document.createElement('div');
  root.className = 'bb-eqg';

  const head = document.createElement('div');
  head.className = 'bb-eqg-head';
  head.textContent = 'Công thức có sẵn';
  root.append(head);

  const list = document.createElement('div');
  list.className = 'bb-eqg-list';
  for (const item of options.items ?? BUILT_IN_EQUATIONS) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'bb-eqg-item';
    const prev = document.createElement('div');
    prev.className = 'bb-eqg-prev';
    prev.textContent = astToLinear(item.ast);
    const name = document.createElement('div');
    name.className = 'bb-eqg-name';
    name.textContent = item.name;
    btn.append(prev, name);
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
  label.textContent = 'Chèn công thức mới';
  neu.append(label);
  if (options.newShortcut) {
    const key = document.createElement('span');
    key.className = 'bb-eqg-key';
    key.textContent = options.newShortcut;
    neu.append(key);
  }
  neu.addEventListener('click', () => options.onNew());
  foot.append(neu);
  root.append(foot);

  return root;
}
