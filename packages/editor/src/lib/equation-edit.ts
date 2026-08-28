/**
 * Pure AST edits for the equation slot editor. Every function returns a NEW
 * tree (the AST rides a node attr — ProseMirror wants immutable updates),
 * addressed by slot paths as the layout emits them: child index / row name,
 * alternating, ending at a row.
 */
import {
  eqCellIndex,
  eqRowNames,
  eqVerticalRows,
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

/** One named row of a node. Matrix cells live in a flat `cells` array rather
 *  than in fields, so their names are read back to an index here — the one
 *  place that knows, so paths stay the plain index/name pairs everything
 *  else walks. */
function rowField(node: unknown, name: string): EqNode[] | null {
  const n = node as Record<string, unknown> | undefined;
  if (!n) return null;
  const cell = eqCellIndex(name);
  if (cell !== null && Array.isArray(n['cells'])) {
    const row = (n['cells'] as EqNode[][])[cell];
    return Array.isArray(row) ? row : null;
  }
  const v = n[name];
  return Array.isArray(v) ? (v as EqNode[]) : null;
}

/** `node` with one named row replaced — the write side of rowField. */
function withRowField(node: EqNode, name: string, row: EqNode[]): EqNode {
  const n = node as unknown as Record<string, unknown>;
  const cell = eqCellIndex(name);
  if (cell !== null && Array.isArray(n['cells']))
    return {
      ...n,
      cells: (n['cells'] as EqNode[][]).map((c, i) => (i === cell ? row : c)),
    } as unknown as EqNode;
  return { ...n, [name]: row } as unknown as EqNode;
}

/** The row a slot path addresses, or null when the path no longer fits. */
export function rowAt(ast: EqNode[], path: Path): EqNode[] | null {
  let row: EqNode[] = ast;
  for (let i = 0; i < path.length; i += 2) {
    const idx = path[i];
    const name = path[i + 1];
    if (typeof idx !== 'number' || typeof name !== 'string') return null;
    const next = rowField(row[idx], name);
    if (!next) return null;
    row = next;
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
    const inner = rowField(n, name);
    if (!inner) return n;
    return withRowField(n, name, withRow(inner, path.slice(2), fn));
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

/**
 * A row that SPELLS something rather than holding maths to edit: a built-in
 * function's name (`sin`), or the word a limit is built on (`lim`). The caret
 * does not stop in one — it goes straight to whatever editable slot the name
 * carries, or past it into the argument. A click still lands there, so a
 * mistyped name is still fixable; only the arrow walk skips it.
 *
 * An EMPTY name is the opposite: that is the Custom entry, where typing the
 * name is the whole point.
 */
function transparent(
  ast: EqNode[],
  path: readonly (number | string)[],
): boolean {
  const row = rowAt(ast, path);
  if (!row || row.length === 0) return false;
  if (path[path.length - 1] === 'name') return true;
  // Inside a name, a row of plain characters is part of the spelling — the
  // `lim` of a limit, the `log` of a logarithm with a base.
  return path.slice(0, -1).includes('name') && row.every((n) => n.t === 'chr');
}

/**
 * The slots of `node` the caret visits, as full paths, in caret order. Minus
 * any the layout did not draw, and seeing THROUGH a transparent row: what it
 * carries is visited in its place, so `lim` offers its condition and then its
 * argument, with nothing in between.
 */
function walkPaths(
  ast: EqNode[],
  slots: readonly EqSlotRect[],
  base: readonly (number | string)[],
  index: number,
  node: EqNode,
): (readonly (number | string)[])[] {
  const out: (readonly (number | string)[])[] = [];
  for (const name of eqRowNames(node)) {
    const path = [...base, index, name];
    if (slotAt(slots, path) < 0) continue;
    if (!transparent(ast, path)) {
      out.push(path);
      continue;
    }
    (rowAt(ast, path) ?? []).forEach((child, i) => {
      if (child.t !== 'chr') out.push(...walkPaths(ast, slots, path, i, child));
    });
  }
  return out;
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
  let from = path;
  // Climbs THROUGH a transparent row rather than stopping in it.
  for (;;) {
    if (from.length < 2) return edge;
    const parent = from.slice(0, -2);
    const index = from[from.length - 2] as number;
    const node = (rowAt(ast, parent) ?? [])[index];
    if (!node) return edge;
    const list = walkPaths(ast, slots, parent, index, node);
    let at = list.findIndex((q) => key(q) === key(from));
    if (at < 0) {
      // Came up out of a transparent row: resume from the outermost slot it
      // contributed, so the next step leaves the whole name behind.
      const mine = list
        .map((q, i) => (isPrefix(from, q) ? i : -1))
        .filter((i) => i >= 0);
      at =
        mine.length === 0
          ? -1
          : dir > 0
            ? Math.max(...mine)
            : Math.min(...mine);
    }
    if (at >= 0) {
      const next = at + dir;
      if (next >= 0 && next < list.length) {
        const i = slotAt(slots, list[next]);
        if (i >= 0)
          return { slot: i, caret: dir > 0 ? 0 : lastCaret(slots[i]) };
      }
    }
    if (!transparent(ast, parent)) {
      const up = slotAt(slots, parent);
      if (up < 0) return edge;
      return { slot: up, caret: dir > 0 ? index + 1 : index };
    }
    from = parent;
  }
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
    const list = walkPaths(ast, slots, cur.path, at, node);
    const p = dir > 0 ? list[0] : list[list.length - 1];
    const i = p ? slotAt(slots, p) : -1;
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
/** The caret stop in `s` nearest to the absolute x the caret came from. */
function nearestCaret(s: EqSlotRect, x: number): number {
  let out = 0;
  let nearest = Infinity;
  s.caretXs.forEach((cx, i) => {
    const d = Math.abs(s.x + cx - x);
    if (d < nearest) {
      nearest = d;
      out = i;
    }
  });
  return out;
}

export function verticalStep(
  ast: EqNode[],
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

  // A row that belongs to its structure's vertical stack steps along that
  // stack, one entry at a time — the structure's own order beats geometry,
  // so a sum's upper limit goes to its LOWER limit and not to the operand
  // sitting nearer below.
  let inStack = false;
  let stackNames: readonly string[] = [];
  let familyPath: readonly (string | number)[] = [];
  if (cur.path.length >= 2) {
    const parent = cur.path.slice(0, -2);
    const index = cur.path[cur.path.length - 2] as number;
    const rowName = cur.path[cur.path.length - 1] as string;
    const node = (rowAt(ast, parent) ?? [])[index];
    if (node) {
      const stack = eqVerticalRows(node).filter(
        (n) =>
          slotAt(slots, [...parent, index, n]) >= 0 &&
          !transparent(ast, [...parent, index, n]),
      );
      const at = stack.indexOf(rowName);
      if (at >= 0) {
        inStack = true;
        stackNames = stack;
        familyPath = [...parent, index];
        const next = at + dir;
        const i =
          next >= 0 && next < stack.length
            ? slotAt(slots, [...parent, index, stack[next]])
            : -1;
        if (i >= 0) return { slot: i, caret: nearestCaret(slots[i], x) };
      }
    }
  }

  // Off the end of the stack — or never on it. Rows the stack already ruled
  // on are out (leaving a denominator would otherwise land in the numerator
  // it just came from), but a sibling OUTSIDE the stack is still fair game:
  // that is how a lone superscript reaches its base, and a sum carrying only
  // an upper limit reaches its operand.
  const ruled = (s: EqSlotRect): boolean =>
    inStack &&
    key(s.path.slice(0, -1)) === key(familyPath) &&
    stackNames.includes(String(s.path[s.path.length - 1]));
  let best = inStack ? -1 : pick(sibling);
  if (best < 0) best = pick((s) => !ruled(s));
  if (best < 0) return null;
  return { slot: best, caret: nearestCaret(slots[best], x) };
}
