/**
 * The document commands — everything an agent can do to ONE document,
 * as data (see ./catalog). Tool names, schemas, anchoring/versioning
 * semantics, and the error texts that teach the retry are defined once
 * here; {@link createMcpServer} and any other adapter read these records.
 */
import { z } from 'zod';
import type { SessionProvider } from './contract.js';
import {
  defineCommand,
  errorText,
  json,
  withSession,
  type AgentCommand,
} from './catalog.js';

const documentId = z
  .string()
  .optional()
  .describe(
    'Target document id. Omit for the currently open document (desktop).',
  );

const expectedVersion = z
  .string()
  .optional()
  .describe(
    'docVersion you last read. If the document changed since, the call fails and you must re-read.',
  );

const occurrence = z
  .number()
  .int()
  .min(1)
  .optional()
  .describe(
    '1-based pick when the anchor text matches more than once (document order).',
  );

export const getDocument = defineCommand({
  name: 'get_document',
  title: 'Read the document',
  description:
    'Read the whole document as numbered blocks (paragraphs, headings — table-cell paragraphs included, in reading order). ' +
    'Returns docVersion: pass it as expectedVersion to mutation tools so concurrent edits are detected. ' +
    'Block indexes are only stable within one docVersion.',
  input: { documentId },
  effect: 'read',
  targets: (a) => [a.documentId],
  run: (provider, { documentId: id }) =>
    withSession(provider, id, async (s) => json(await s.snapshot())),
});

export const findText = defineCommand({
  name: 'find_text',
  title: 'Find text',
  description:
    'Find every occurrence of a text in the document. Matches are within one paragraph (they never span paragraphs or inline objects). ' +
    'Returns each match with its block index, 1-based occurrence number, and surrounding context.',
  input: {
    documentId,
    query: z.string().min(1).describe('Exact text to find (case-sensitive).'),
  },
  effect: 'read',
  targets: (a) => [a.documentId],
  run: (provider, { documentId: id, query }) =>
    withSession(provider, id, async (s) =>
      json({ matches: await s.find(query) }),
    ),
});

export const replaceText = defineCommand({
  name: 'replace_text',
  title: 'Replace text',
  description:
    'Replace one occurrence of exact text. old_text must match exactly once in the document — if it matches more, ' +
    'either pass occurrence or use a longer, unique anchor (include surrounding words). Formatting of the replaced range is kept.',
  input: {
    documentId,
    old_text: z
      .string()
      .min(1)
      .describe(
        'Exact existing text (must match uniquely, or pass occurrence).',
      ),
    new_text: z.string().describe('Replacement text (empty string deletes).'),
    occurrence,
    expectedVersion,
  },
  effect: 'edit',
  targets: (a) => [a.documentId],
  run: (
    provider,
    {
      documentId: id,
      old_text,
      new_text,
      occurrence: occ,
      expectedVersion: ver,
    },
  ) =>
    withSession(provider, id, async (s) =>
      json(
        await s.replaceText(old_text, new_text, {
          occurrence: occ,
          expectedVersion: ver,
        }),
      ),
    ),
});

export const insertContent = defineCommand({
  name: 'insert_content',
  title: 'Insert content',
  description:
    'Insert new paragraphs. Each line of `content` becomes one paragraph. Anchor by exact text with position before/after ' +
    '(the paragraph containing the anchor), or position document_end to append.',
  input: {
    documentId,
    content: z
      .string()
      .min(1)
      .describe('Plain text; every line becomes a paragraph.'),
    position: z.enum(['before', 'after', 'document_end']),
    anchor_text: z
      .string()
      .optional()
      .describe(
        'Required for before/after: exact text inside the anchor paragraph.',
      ),
    occurrence,
    expectedVersion,
  },
  effect: 'edit',
  targets: (a) => [a.documentId],
  run: (
    provider,
    {
      documentId: id,
      content,
      position,
      anchor_text,
      occurrence: occ,
      expectedVersion: ver,
    },
  ) =>
    withSession(provider, id, async (s) => {
      if (position === 'document_end') {
        return json(
          await s.insertContent(
            content,
            { position: 'document_end' },
            { expectedVersion: ver },
          ),
        );
      }
      if (!anchor_text)
        return errorText(
          'anchor_text is required when position is before/after.',
        );
      return json(
        await s.insertContent(
          content,
          { position, text: anchor_text, occurrence: occ },
          { expectedVersion: ver },
        ),
      );
    }),
});

export const applyFormatting = defineCommand({
  name: 'apply_formatting',
  title: 'Apply formatting',
  description:
    'Format one occurrence of exact text: character marks (bold/italic/underline/strike — true applies, false removes) ' +
    'and/or paragraph alignment of the containing paragraph. Same anchoring rules as replace_text.',
  input: {
    documentId,
    target_text: z
      .string()
      .min(1)
      .describe(
        'Exact text to format (must match uniquely, or pass occurrence).',
      ),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    strike: z.boolean().optional(),
    align: z.enum(['left', 'center', 'right', 'justify']).optional(),
    occurrence,
    expectedVersion,
  },
  effect: 'edit',
  targets: (a) => [a.documentId],
  run: (
    provider,
    {
      documentId: id,
      target_text,
      occurrence: occ,
      expectedVersion: ver,
      ...format
    },
  ) =>
    withSession(provider, id, async (s) =>
      json(
        await s.applyFormatting(target_text, format, {
          occurrence: occ,
          expectedVersion: ver,
        }),
      ),
    ),
});

export const updateImage = defineCommand({
  name: 'update_image',
  title: 'Resize / rotate an image',
  description:
    'Resize and/or rotate one image. Read get_document first: blocks list their images, and ' +
    '(block_index, image_index) addresses one. Omitted fields keep their current value; width/height are CSS px, ' +
    'rotation is clockwise degrees around the image center (normalized to 0-360; 0 = upright). ' +
    'Block indexes change with every edit — pass expectedVersion from your last read.',
  input: {
    documentId,
    block_index: z
      .number()
      .int()
      .min(0)
      .describe('Index of the block holding the image (from get_document).'),
    image_index: z
      .number()
      .int()
      .min(0)
      .optional()
      .describe('0-based image within the block (default 0, the first).'),
    width: z.number().positive().optional().describe('New width in CSS px.'),
    height: z.number().positive().optional().describe('New height in CSS px.'),
    rotation: z
      .number()
      .optional()
      .describe('Clockwise degrees around the center.'),
    expectedVersion,
  },
  effect: 'edit',
  targets: (a) => [a.documentId],
  run: (
    provider,
    {
      documentId: id,
      block_index,
      image_index,
      width,
      height,
      rotation,
      expectedVersion: ver,
    },
  ) =>
    withSession(provider, id, async (s) => {
      if (
        width === undefined &&
        height === undefined &&
        rotation === undefined
      ) {
        return errorText('Pass at least one of width, height, rotation.');
      }
      return json(
        await s.updateImage(
          block_index,
          image_index ?? 0,
          { width, height, rotation },
          { expectedVersion: ver },
        ),
      );
    }),
});

export const saveDocument = defineCommand({
  name: 'save_document',
  title: 'Save the document',
  description:
    'Persist the document to its backing store (file on desktop). Usually optional — the desktop autosaves.',
  input: { documentId },
  effect: 'save',
  targets: (a) => [a.documentId],
  run: (provider, { documentId: id }) =>
    withSession(provider, id, async (s) => {
      await s.save();
      return json({ saved: true });
    }),
});

export const getSelection = defineCommand({
  name: 'get_selection',
  title: "Read the user's selection",
  description:
    'The text the user currently has selected in the editor, or null when nothing is selected.',
  input: { documentId },
  effect: 'read',
  requires: 'selection',
  targets: (a) => [a.documentId],
  run: (provider, { documentId: id }) =>
    withSession(provider, id, async (s) =>
      json((await s.getSelection?.()) ?? { selection: null }),
    ),
});

/** Every document command, in the order an agent reads them. `get_selection`
 *  needs the `selection` capability and is left out by hosts without one. */
export const documentCommands: readonly AgentCommand<
  z.ZodRawShape,
  SessionProvider
>[] = [
  getDocument,
  findText,
  replaceText,
  insertContent,
  applyFormatting,
  updateImage,
  saveDocument,
  getSelection,
];
