/**
 * Pure AST edits for the equation slot editor. Every function returns a NEW
 * tree (the AST rides a node attr — ProseMirror wants immutable updates),
 * addressed by slot paths as the layout emits them: child index / row name,
 * alternating, ending at a row.
 */
import {
  eqRowNames,
  MATH_ALPHABETS,
  MATH_AUTOCORRECT,
  mathLetters,
  type EqNode,
  type EqSlotRect,
} from '@shadow-garden/bapbong-contracts';

type Path = readonly (number | string)[];

/** A math-italic letter back to its ASCII spelling (typed letters are
 *  letterformed on insert, so a `\name` run is italic in the tree). */
function unstyled(ch: string): string {
  const cp = ch.codePointAt(0) ?? 0;
  if (cp >= 0x1d434 && cp <= 0x1d44d)
    return String.fromCharCode(65 + (cp - 0x1d434));
  if (cp >= 0x1d44e && cp <= 0x1d467)
    return String.fromCharCode(97 + (cp - 0x1d44e));
  if (cp === 0x210e) return 'h';
  return ch;
}

/** The row a slot path addresses, or null when the path no longer fits. */
export function rowAt(ast: EqNode[], path: Path): EqNode[] | null {
  let row: EqNode[] = ast;
  for (let i = 0; i < path.length; i += 2) {
    const idx = path[i];
    const name = path[i + 1];
    if (typeof idx !== 'number' || typeof name !== 'string') return null;
    const node = row[idx] as unknown as Record<string, unknown> | undefined;
    const next = node?.[name];
    if (!Array.isArray(next)) return null;
    row = next as EqNode[];
  }
  return row;
}

/** The AST with the row at `path` replaced by `fn(row)`. */
export function withRow(
  ast: EqNode[],
  path: Path,
  fn: (row: EqNode[]) => EqNode[],
): EqNode[] {
  if (path.length === 0) return fn(ast);
  const idx = path[0] as number;
  const name = path[1] as string;
  return ast.map((n, i) => {
    if (i !== idx) return n;
    const node = n as unknown as Record<string, unknown>;
    return {
      ...node,
      [name]: withRow(node[name] as EqNode[], path.slice(2), fn),
    } as unknown as EqNode;
  });
}

/** A typed character, letterformed the way every equation spells letters. */
export function eqChar(ch: string): EqNode {
  return { t: 'chr', ch: mathLetters(ch, MATH_ALPHABETS['italic']) };
}

/** Insert one item at `caret` in the row at `path`. */
export function insertAt(
  ast: EqNode[],
  path: Path,
  caret: number,
  node: EqNode,
): EqNode[] {
  return withRow(ast, path, (row) => [
    ...row.slice(0, caret),
    node,
    ...row.slice(caret),
  ]);
}

/** Remove the item before `caret` in the row at `path` (no-op at 0). */
export function removeBefore(
  ast: EqNode[],
  path: Path,
  caret: number,
): EqNode[] {
  if (caret <= 0) return ast;
  return withRow(ast, path, (row) => [
    ...row.slice(0, caret - 1),
    ...row.slice(caret),
  ]);
}

/**
 * The Math AutoCorrect a trigger space completes at `caret`: the `\name`
 * spelled by the chr items right before it. Returns the replaced row and
 * the caret after the symbol, or null when nothing matches.
 */
export function autoCorrectAt(
  ast: EqNode[],
  path: Path,
  caret: number,
): { ast: EqNode[]; caret: number } | null {
  const row = rowAt(ast, path);
  if (!row) return null;
  let start = caret;
  let name = '';
  while (start > 0) {
    const n = row[start - 1];
    if (n.t !== 'chr') break;
    const plain = unstyled(n.ch);
    if (!/^[A-Za-z\\]$/.test(plain)) break;
    name = plain + name;
    start--;
    if (plain === '\\') break;
  }
  if (!name.startsWith('\\')) return null;
  const to = MATH_AUTOCORRECT[name.slice(1)];
  if (!to) return null;
  const next = withRow(ast, path, (r) => [
    ...r.slice(0, start),
    ...[...to].map((ch) => ({ t: 'chr' as const, ch })),
    ...r.slice(caret),
  ]);
  return { ast: next, caret: start + [...to].length };
}

/** Structures the editor builds from a keystroke inside a slot. */
export function structureFor(key: string): EqNode | null {
  if (key === '/') return { t: 'frac', num: [], den: [] };
  if (key === '^')
    return { t: 'scr', base: [], sub: [], sup: [], slots: 'sup' };
  if (key === '_')
    return { t: 'scr', base: [], sub: [], sup: [], slots: 'sub' };
  return null;
}

// ── Caret walking ─────────────────────────────────────────────────────
//
// The layout gives every editable row a rect with the caret X for each
// boundary in it. A row treats a child structure as ONE step, which is right
// for drawing and wrong for editing: stepping right past a fraction should go
// INTO its numerator, not over the whole thing. These functions turn the flat
// slot list back into the tree it came from (each rect carries its path) and
// walk it.

/** Where a step landed: a slot and a caret in it, the equation's left or
 *  right edge, or null for "not mine — let the document move the caret". */
export type SlotStep =
  | { slot: number; caret: number }
  | 'out-left'
  | 'out-right'
  | null;

const key = (p: readonly (number | string)[]): string => JSON.stringify(p);

const isPrefix = (
  a: readonly (number | string)[],
  b: readonly (number | string)[],
): boolean => a.length < b.length && a.every((v, i) => v === b[i]);

function slotAt(
  slots: readonly EqSlotRect[],
  path: readonly (number | string)[],
): number {
  const k = key(path);
  return slots.findIndex((s) => key(s.path) === k);
}

/** The rows of `node` in caret order, minus any the layout did not draw. */
function rowsOf(
  node: EqNode,
  base: readonly (number | string)[],
  index: number,
  slots: readonly EqSlotRect[],
): string[] {
  return eqRowNames(node).filter(
    (n) => slotAt(slots, [...base, index, n]) >= 0,
  );
}

const lastCaret = (s: EqSlotRect): number => Math.max(0, s.caretXs.length - 1);

/** Out of the row at `path` in direction `dir`: the next row of the same
 *  structure (numerator → denominator), or back up to the parent, landing
 *  just past the structure we were inside. */
function ascend(
  ast: EqNode[],
  slots: readonly EqSlotRect[],
  path: readonly (number | string)[],
  dir: 1 | -1,
): SlotStep {
  const edge: SlotStep = dir > 0 ? 'out-right' : 'out-left';
  if (path.length < 2) return edge;
  const parent = path.slice(0, -2);
  const index = path[path.length - 2] as number;
  const rowName = path[path.length - 1] as string;
  const node = (rowAt(ast, parent) ?? [])[index];
  if (!node) return edge;
  const names = rowsOf(node, parent, index, slots);
  const next = names.indexOf(rowName) + dir;
  if (next >= 0 && next < names.length) {
    const i = slotAt(slots, [...parent, index, names[next]]);
    if (i >= 0) return { slot: i, caret: dir > 0 ? 0 : lastCaret(slots[i]) };
  }
  const up = slotAt(slots, parent);
  if (up < 0) return edge;
  return { slot: up, caret: dir > 0 ? index + 1 : index };
}

/** One step left or right, descending into structures and climbing out of
 *  them — the caret visits every position a user can type at, in reading
 *  order. */
export function horizontalStep(
  ast: EqNode[],
  slots: readonly EqSlotRect[],
  slot: number,
  caret: number,
  dir: 1 | -1,
): SlotStep {
  const cur = slots[slot];
  if (!cur) return null;
  const row = rowAt(ast, cur.path) ?? [];
  const at = dir > 0 ? caret : caret - 1;
  const node = row[at];
  if (dir > 0 ? caret >= row.length : caret <= 0)
    return ascend(ast, slots, cur.path, dir);
  if (node && node.t !== 'chr') {
    const names = rowsOf(node, cur.path, at, slots);
    const name = dir > 0 ? names[0] : names[names.length - 1];
    const i = name ? slotAt(slots, [...cur.path, at, name]) : -1;
    if (i >= 0) return { slot: i, caret: dir > 0 ? 0 : lastCaret(slots[i]) };
  }
  return { slot, caret: caret + dir };
}

/** Straight up or down to the nearest row on that side, keeping the caret's
 *  horizontal place — the shortcut from a numerator to its denominator
 *  without walking to the end of it first.
 *
 *  Ancestors and descendants are not candidates: the row you are in is drawn
 *  inside its parent's box, so "the row below" would otherwise always find
 *  the enclosing row it is already part of. */
export function verticalStep(
  slots: readonly EqSlotRect[],
  slot: number,
  caret: number,
  dir: 1 | -1,
): SlotStep {
  const cur = slots[slot];
  if (!cur) return null;
  const x = cur.x + (cur.caretXs[Math.min(caret, lastCaret(cur))] ?? 0);
  const from = cur.y + cur.height / 2;
  // Rows of the SAME structure — a fraction's other half, an operator's
  // other limit. They share everything but the row name.
  const family = key(cur.path.slice(0, -1));
  const sibling = (s: EqSlotRect): boolean =>
    cur.path.length >= 2 &&
    s.path.length === cur.path.length &&
    key(s.path.slice(0, -1)) === family;

  const pick = (only: (s: EqSlotRect) => boolean): number => {
    let best = -1;
    let bestScore = Infinity;
    slots.forEach((s, i) => {
      if (
        i === slot ||
        isPrefix(s.path, cur.path) ||
        isPrefix(cur.path, s.path)
      )
        return;
      if (!only(s)) return;
      const mid = s.y + s.height / 2;
      if (dir > 0 ? mid <= from + 0.5 : mid >= from - 0.5) return;
      const dx =
        x < s.x ? s.x - x : x > s.x + s.width ? x - (s.x + s.width) : 0;
      // Vertical nearness decides; horizontal distance only breaks ties, so a
      // row directly below wins over one that is closer across the page.
      const score = Math.abs(mid - from) * 4 + dx;
      if (score < bestScore) {
        bestScore = score;
        best = i;
      }
    });
    return best;
  };

  // Inside a structure, up and down mean the rows of THAT structure: from a
  // sum's upper limit, down is its operand — never a bracket that happens to
  // sit closer across the page. Only when the structure has nothing that way
  // does the search widen to the whole equation.
  let best = pick(sibling);
  if (best < 0) best = pick(() => true);
  if (best < 0) return null;
  const s = slots[best];
  let caretOut = 0;
  let nearest = Infinity;
  s.caretXs.forEach((cx, i) => {
    const d = Math.abs(s.x + cx - x);
    if (d < nearest) {
      nearest = d;
      caretOut = i;
    }
  });
  return { slot: best, caret: caretOut };
}
