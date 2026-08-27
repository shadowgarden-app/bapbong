/**
 * Pure AST edits for the equation slot editor. Every function returns a NEW
 * tree (the AST rides a node attr — ProseMirror wants immutable updates),
 * addressed by slot paths as the layout emits them: child index / row name,
 * alternating, ending at a row.
 */
import {
  MATH_ALPHABETS,
  MATH_AUTOCORRECT,
  mathLetters,
  type EqNode,
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
