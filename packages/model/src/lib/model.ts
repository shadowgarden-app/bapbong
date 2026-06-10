import { Schema } from 'prosemirror-model';

/**
 * bapbong's ProseMirror document schema.
 *
 * Deliberately minimal for M1 (DOCX import vertical slice): block paragraphs
 * of inline text with the four common character toggles. Lists, tables,
 * headings, images, etc. are added in later milestones — when they land they
 * extend THIS schema so the importer, layout engine, and (canvas) painter all
 * agree on one document model.
 */
export const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },

    paragraph: {
      group: 'block',
      content: 'inline*',
      parseDOM: [{ tag: 'p' }],
      toDOM: () => ['p', 0],
    },

    text: { group: 'inline' },
  },

  marks: {
    // w:b
    strong: {
      parseDOM: [{ tag: 'strong' }, { tag: 'b' }],
      toDOM: () => ['strong', 0],
    },
    // w:i
    em: {
      parseDOM: [{ tag: 'em' }, { tag: 'i' }],
      toDOM: () => ['em', 0],
    },
    // w:u
    underline: {
      parseDOM: [{ tag: 'u' }],
      toDOM: () => ['u', 0],
    },
    // w:strike
    strike: {
      parseDOM: [{ tag: 's' }, { tag: 'strike' }],
      toDOM: () => ['s', 0],
    },
  },
});

/** Concrete schema type, handy for typing Node/Mark across packages. */
export type BapbongSchema = typeof schema;
