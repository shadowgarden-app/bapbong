/**
 * Footnote bodies written in bapbong.
 *
 * An imported footnote's body lives outside the document (the importer hands
 * it to the layout beside the doc, keyed by its reference's `num`); a
 * footnote made here has nowhere else to live, so its body rides the doc as
 * `doc.attrs.footnoteBodies` (num → story JSON), undoable with the rest. The
 * layout, the editor and an agent reading the document combine the two the
 * same way — this is that one place.
 */
import type { Node as PMNode } from 'prosemirror-model';

/** num → the number the reference shows (its text), in document order. */
export function footnoteNumbers(doc: PMNode): Map<number, string> {
  const out = new Map<number, string>();
  doc.descendants((n) => {
    if (!n.isText) return true;
    const fn = n.marks.find((m) => m.type.name === 'footnote');
    const num = fn ? Number(fn.attrs['num']) : NaN;
    if (Number.isFinite(num) && !out.has(num)) out.set(num, n.text ?? '');
    return false;
  });
  return out;
}

/**
 * Every footnote body the layout shows: the imported ones, plus the ones in
 * `doc.attrs.footnoteBodies` — each of those prefixed with its reference's
 * number as a superscript, the way the importer prefixes imported bodies.
 * `cache` (keyed by the attr's JSON objects) keeps a parse per body.
 */
export function effectiveFootnotes(
  doc: PMNode,
  imported: Record<number, PMNode> | undefined,
  cache?: WeakMap<object, PMNode>,
): Record<number, PMNode> | undefined {
  const bodies = doc.attrs['footnoteBodies'] as Record<string, unknown> | null;
  if (!bodies || Object.keys(bodies).length === 0) return imported;
  const schema = doc.type.schema;
  const numbers = footnoteNumbers(doc);
  const out: Record<number, PMNode> = { ...(imported ?? {}) };
  for (const [key, json] of Object.entries(bodies)) {
    const num = Number(key);
    if (!Number.isFinite(num) || json == null || typeof json !== 'object')
      continue;
    let story = cache?.get(json);
    if (!story) {
      try {
        story = schema.nodeFromJSON(json);
      } catch {
        continue; // malformed attr data must not take the layout down
      }
      cache?.set(json, story);
    }
    const display = numbers.get(num);
    out[num] = display ? withMarker(story, display) : story;
  }
  return out;
}

function withMarker(story: PMNode, display: string): PMNode {
  const schema = story.type.schema;
  const sup = schema.marks['vertAlign']?.create({ value: 'super' });
  const marker = schema.text(display, sup ? [sup] : []);
  const first = story.firstChild;
  const blocks: PMNode[] = [];
  story.forEach((b) => blocks.push(b));
  if (first && first.type.name === 'paragraph') {
    const kids: PMNode[] = [marker, schema.text(' ')];
    first.forEach((k) => kids.push(k));
    blocks[0] = first.type.create(first.attrs, kids, first.marks);
  } else {
    blocks.unshift(schema.nodes['paragraph'].create(null, [marker]));
  }
  return story.type.create(story.attrs, blocks);
}
