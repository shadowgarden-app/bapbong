/**
 * Bookmarks and generated fields — the two things that make a table of
 * contents behave like Word's rather than like a list of ordinary links.
 *
 * A `w:bookmarkStart` names a place in the document; a hyperlink whose href
 * is `#name` points at it. Word's TOC wires every entry this way, with
 * machine-generated `_Toc…` names it never shows the reader. Bookmarks ride
 * the paragraph that contains them (see model.ts `bookmarks`), so they move
 * with their heading through edits instead of decaying into stale offsets.
 */

import type { Node as PMNode } from 'prosemirror-model';

/** A generated field's identity, as it rides a paragraph's `field` attr. */
export interface FieldInfo {
  /** Field kind — only 'toc' is modelled today. */
  kind: string;
  /** The field instruction, e.g. `TOC \o "1-3" \h \z \u`. */
  instr: string;
}

/** An href pointing inside this document (`#name`) → the bookmark name, or
 *  null for external/absent links. */
export function anchorName(href: string | null | undefined): string | null {
  return href && href.startsWith('#') && href.length > 1 ? href.slice(1) : null;
}

/** Position of the paragraph anchoring `name`, or null when no paragraph
 *  claims it (a stale link, or a bookmark outside a paragraph). The position
 *  is the paragraph's first content slot — where a caret should land. */
export function findBookmark(doc: PMNode, name: string): number | null {
  let found: number | null = null;
  doc.descendants((node, pos) => {
    if (found !== null) return false;
    if (node.type.name !== 'paragraph') return true; // tables hold paragraphs
    const names = node.attrs['bookmarks'] as string[] | null;
    if (names?.includes(name)) found = pos + 1;
    return false;
  });
  return found;
}

/** The text a reader should see for an internal link — the target paragraph's
 *  own text, trimmed to `max` — or null when the target is missing. Beats the
 *  raw `_Toc89595219`, which is machine bookkeeping. */
export function bookmarkLabel(
  doc: PMNode,
  name: string,
  max = 60,
): string | null {
  const pos = findBookmark(doc, name);
  if (pos === null) return null;
  const para = doc.nodeAt(pos - 1);
  // A numbered heading's marker ("CHƯƠNG 1") is generated at layout time and
  // isn't part of the text, so the raw content can start mid-punctuation
  // (": GIỚI THIỆU") — lead with the actual words instead.
  const text = (para?.textContent ?? '').replace(/^[\s.:–—-]+/, '').trim();
  if (!text) return null;
  return text.length > max ? `${text.slice(0, max - 1)}…` : text;
}

/** The field a position sits inside, with the full paragraph range it spans,
 *  or null. A field covers CONSECUTIVE paragraphs carrying the same `field`
 *  attr object — the importer stamps one shared object per field, so identity
 *  distinguishes two adjacent TOCs. */
export function fieldAt(
  doc: PMNode,
  pos: number,
): { field: FieldInfo; from: number; to: number } | null {
  type Run = { field: FieldInfo; from: number; to: number };
  let hit: Run | null = null;
  let run: Run | null = null;
  // Each run must be walked to its END before it is returned — a caret in the
  // field's FIRST paragraph still belongs to the whole span, so this can't
  // stop at the paragraph that contains `pos`.
  doc.forEach((node, offset) => {
    const f = node.attrs['field'] as FieldInfo | null;
    if (f && run && run.field === f) {
      run.to = offset + node.nodeSize;
    } else {
      if (run && hit === run) return; // the matched run just ended — keep it
      run = f ? { field: f, from: offset, to: offset + node.nodeSize } : null;
    }
    if (run && hit === null && pos >= run.from && pos <= run.to) hit = run;
  });
  return hit;
}
