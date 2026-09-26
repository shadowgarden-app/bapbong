/**
 * The document commands — everything an agent can do to ONE document,
 * as data (see ./catalog). Tool names, schemas, anchoring/versioning
 * semantics, and the error texts that teach the retry are defined once
 * here; {@link createMcpServer} and any other adapter read these records.
 */
import { z } from 'zod';
import { ContentError, type SessionProvider } from './contract.js';
import {
  defineCommand,
  errorText,
  json,
  withSession,
  type AgentCommand,
} from './catalog.js';
import {
  CONTENT_GRAMMAR,
  contentSchema,
  tableEditShape,
  tabStopSchema,
} from './blocks-schema.js';

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

const lengthArg = z
  .union([z.number(), z.string()])
  .describe('cm as a number, or "2cm", "1in", "25mm".');

/** "1.15" → a multiple; "18pt" → exact; "at least 12pt" → a minimum. */
function lineSpacingOf(
  v: number | string,
): number | { exact: number } | { atLeast: number } {
  if (typeof v === 'number') return v;
  const t = v.trim();
  const least = /^at\s*least\s+(.+)$/i.exec(t);
  const pt = (x: string) => {
    const m = /^(\d+(?:\.\d+)?)\s*pt$/i.exec(x.trim());
    if (!m)
      throw new ContentError(
        `line_spacing ${JSON.stringify(v)}: a height is in points, like "18pt".`,
      );
    return Number(m[1]);
  };
  if (least) return { atLeast: pt(least[1]) };
  if (/^\d+(\.\d+)?$/.test(t)) return Number(t);
  return { exact: pt(t) };
}

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
    'A list item says so (list: { kind, level }); its bullet or number is drawn, not in its text. ' +
    'Hyperlinks are listed per block (links: [{ text, href }]). Headers and footers with text come as chrome ' +
    '(part, variant, the sections showing it, text; {page} is a page number) — edit them with edit_header_footer. ' +
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
    'Insert content: paragraphs, headings, tables. Anchor by exact text with position before/after ' +
    '(the paragraph containing the anchor), or position document_end to append. ' +
    CONTENT_GRAMMAR,
  input: {
    documentId,
    content: contentSchema.describe(
      'Plain text (one paragraph per line) or an array of blocks — see the tool description.',
    ),
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
    'Format text or a paragraph. Address it by exact text (target_text, matched once or with occurrence — same rules ' +
    'as replace_text) or a whole block (block_index from get_document). Character marks bold/italic/underline/strike ' +
    '(true applies, false removes), font_size, font, color, highlight, vertical_align and link apply to the text ' +
    '(clear_formatting strips the text back to plain first); align, heading (1-6, 0 = body text), style ' +
    "(Title, Subtitle, Normal or any style of the document — list_styles), tabs (replace the paragraph's tab stops), list, space_before/space_after (pt), " +
    'line_spacing, indent_left/indent_right/first_line/hanging (cm or "1cm", 0 removes) apply to the containing paragraph. ' +
    'list "bullet"/"number" makes it a list item — it joins a list of that kind right above it, so turning ' +
    'several paragraphs into one list is one call per paragraph, top to bottom; "none" makes it body text. ' +
    'list_level (1-3) nests a list item.',
  input: {
    documentId,
    target_text: z
      .string()
      .min(1)
      .optional()
      .describe(
        'Exact text to format (must match uniquely, or pass occurrence).',
      ),
    block_index: z
      .number()
      .int()
      .min(0)
      .optional()
      .describe(
        'Instead of target_text: the whole block with this index (from get_document).',
      ),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    strike: z.boolean().optional(),
    font_size: z.number().positive().optional().describe('Points.'),
    font: z
      .string()
      .min(1)
      .max(100)
      .optional()
      .describe('Font family, as Word shows it; "default" removes it.'),
    color: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$|^auto$/)
      .optional()
      .describe('Text colour "#RRGGBB", or "auto".'),
    highlight: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$|^none$/)
      .optional()
      .describe('Highlight colour "#RRGGBB", or "none".'),
    vertical_align: z.enum(['superscript', 'subscript', 'baseline']).optional(),
    link: z
      .string()
      .min(1)
      .optional()
      .describe(
        'Make the text a hyperlink: a web address, "mailto:…", "tel:…" or "#bookmark"; "none" unlinks it.',
      ),
    clear_formatting: z
      .boolean()
      .optional()
      .describe(
        'Strip character formatting first (links and comments stay), then apply the rest of this call.',
      ),
    align: z.enum(['left', 'center', 'right', 'justify']).optional(),
    heading: z
      .number()
      .int()
      .min(0)
      .max(6)
      .optional()
      .describe('Word Heading level; 0 makes it body text.'),
    style: z
      .string()
      .min(1)
      .optional()
      .describe(
        'Title, Subtitle, Normal, or a style the document defines, by name or id (list_styles) — its look comes with it.',
      ),
    tabs: z
      .array(tabStopSchema)
      .optional()
      .describe(
        "The paragraph's tab stops, replacing any it had; [] removes them.",
      ),
    space_before: z
      .number()
      .min(0)
      .optional()
      .describe('Space above the paragraph, points.'),
    space_after: z
      .number()
      .min(0)
      .optional()
      .describe('Space below the paragraph, points.'),
    line_spacing: z
      .union([z.number(), z.string()])
      .optional()
      .describe(
        'A multiple of single (1, 1.15, 1.5, 2), an exact height ("18pt") or a minimum ("at least 12pt").',
      ),
    indent_left: lengthArg.optional(),
    indent_right: lengthArg.optional(),
    first_line: lengthArg
      .optional()
      .describe('First-line indent; clears hanging. 0 removes it.'),
    hanging: lengthArg
      .optional()
      .describe('Hanging indent; clears first_line. 0 removes it.'),
    list: z
      .enum(['bullet', 'number', 'none'])
      .optional()
      .describe('Make the paragraph a list item, or "none" for body text.'),
    list_level: z
      .number()
      .int()
      .min(1)
      .max(9)
      .optional()
      .describe('Nesting level of a list item, 1 = top.'),
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
      block_index,
      occurrence: occ,
      expectedVersion: ver,
      font_size,
      font,
      color,
      highlight,
      vertical_align,
      clear_formatting,
      link,
      heading,
      style,
      list,
      list_level,
      space_before,
      space_after,
      line_spacing,
      indent_left,
      indent_right,
      first_line,
      hanging,
      ...format
    },
  ) =>
    withSession(provider, id, async (s) => {
      if (target_text === undefined && block_index === undefined)
        return errorText(
          'Pass target_text (exact text) or block_index (a whole block).',
        );
      return json(
        await s.applyFormatting(
          target_text !== undefined
            ? target_text
            : { blockIndex: block_index as number },
          {
            ...format,
            ...(font_size !== undefined ? { fontSize: font_size } : {}),
            ...(font !== undefined
              ? { fontFamily: font === 'default' ? null : font }
              : {}),
            ...(color !== undefined
              ? { color: color === 'auto' ? null : color }
              : {}),
            ...(highlight !== undefined
              ? { highlight: highlight === 'none' ? null : highlight }
              : {}),
            ...(vertical_align !== undefined
              ? {
                  verticalAlign:
                    vertical_align === 'baseline' ? null : vertical_align,
                }
              : {}),
            ...(clear_formatting ? { clear: true } : {}),
            ...(link !== undefined
              ? { link: link === 'none' ? null : link }
              : {}),
            ...(heading !== undefined ? { heading } : {}),
            ...(style !== undefined
              ? { style: style === 'Normal' ? null : style }
              : {}),
            ...(list !== undefined
              ? { list: list === 'none' ? null : list }
              : {}),
            ...(list_level !== undefined ? { listLevel: list_level } : {}),
            ...(space_before !== undefined
              ? { spaceBefore: space_before }
              : {}),
            ...(space_after !== undefined ? { spaceAfter: space_after } : {}),
            ...(line_spacing !== undefined
              ? { lineSpacing: lineSpacingOf(line_spacing) }
              : {}),
            ...(indent_left !== undefined ||
            indent_right !== undefined ||
            first_line !== undefined ||
            hanging !== undefined
              ? {
                  indent: {
                    ...(indent_left !== undefined ? { left: indent_left } : {}),
                    ...(indent_right !== undefined
                      ? { right: indent_right }
                      : {}),
                    ...(first_line !== undefined
                      ? { firstLine: first_line }
                      : {}),
                    ...(hanging !== undefined ? { hanging } : {}),
                  },
                }
              : {}),
          },
          { occurrence: occ, expectedVersion: ver },
        ),
      );
    }),
});

export const editTable = defineCommand({
  name: 'edit_table',
  title: 'Edit a table',
  description:
    'Change an existing table: insert or delete rows and columns, merge cells in a row, set column widths, borders ' +
    '(grid/outer/none), a repeated header row, or the table alignment — or remove the whole table with ' +
    'delete_table. Address the table by its 0-based index: get_document marks every block inside a table with ' +
    'table: { index, row, cell }. Rows, columns and cells are 0-based. Inserted columns are empty and the table ' +
    'keeps its width unless you pass widths. Order within one call: delete_rows, delete_columns, insert_columns, ' +
    'insert_rows (rows cover the columns the table has by then), merge, widths, then borders/header/align. ' +
    'Cell text is edited with replace_text like any paragraph.',
  input: {
    documentId,
    table: z
      .number()
      .int()
      .min(0)
      .describe('The table, 0-based, from get_document.'),
    ...tableEditShape,
    expectedVersion,
  },
  effect: 'edit',
  targets: (a) => [a.documentId],
  run: (
    provider,
    {
      documentId: id,
      table,
      delete_table,
      insert_columns,
      delete_columns,
      insert_rows,
      delete_rows,
      merge,
      widths,
      borders,
      header,
      align,
      expectedVersion: ver,
    },
  ) =>
    withSession(provider, id, async (s) => {
      if (
        !delete_table &&
        !insert_columns &&
        !delete_columns?.length &&
        !insert_rows &&
        !delete_rows?.length &&
        !merge &&
        !widths &&
        !borders &&
        header === undefined &&
        align === undefined
      ) {
        return errorText(
          'Pass at least one change: insert_rows, delete_rows, insert_columns, delete_columns, merge, widths, ' +
            'borders, header, align — or delete_table.',
        );
      }
      return json(
        await s.editTable(
          table,
          {
            ...(delete_table ? { deleteTable: true } : {}),
            ...(insert_columns ? { insertColumns: insert_columns } : {}),
            ...(delete_columns?.length
              ? { deleteColumns: delete_columns }
              : {}),
            ...(insert_rows ? { insertRows: insert_rows } : {}),
            ...(delete_rows ? { deleteRows: delete_rows } : {}),
            ...(merge ? { merge } : {}),
            ...(widths ? { widths } : {}),
            ...(borders ? { borders } : {}),
            ...(header !== undefined ? { header } : {}),
            ...(align !== undefined ? { align } : {}),
          },
          { expectedVersion: ver },
        ),
      );
    }),
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

/** The three ways a picture can reach a document, as tool arguments. One of
 *  them, and which ones an app actually accepts is in its answer when it
 *  cannot take the one you passed. */
const imageSourceShape = {
  image_path: z
    .string()
    .optional()
    .describe(
      'Full path of a picture file to use (PNG, JPEG, GIF or BMP). The app reads it only if you may read that folder.',
    ),
  attachment_id: z
    .string()
    .optional()
    .describe(
      'Id of a picture the user attached to this conversation — how you use a picture they sent you.',
    ),
  svg: z
    .string()
    .optional()
    .describe(
      'SVG markup to draw the picture yourself. The app rasterizes it before it goes in.',
    ),
};

const ONE_SOURCE =
  'Give the picture exactly one way: image_path (a file), attachment_id (a picture the user sent) ' +
  'or svg (markup you write). ';

/** Zod cannot say "exactly one of these" in a shape MCP clients read well,
 *  so the check is here, with a sentence the model can act on. */
function imageSource(a: {
  image_path?: string;
  attachment_id?: string;
  svg?: string;
}):
  | { kind: 'path'; path: string }
  | { kind: 'attachment'; attachmentId: string }
  | { kind: 'svg'; svg: string }
  | string {
  const given = [
    a.image_path !== undefined && ('path' as const),
    a.attachment_id !== undefined && ('attachment' as const),
    a.svg !== undefined && ('svg' as const),
  ].filter(Boolean);
  if (given.length !== 1)
    return given.length === 0
      ? `No picture given. ${ONE_SOURCE}`
      : `Pass only one picture, not ${given.length}. ${ONE_SOURCE}`;
  if (a.image_path !== undefined) return { kind: 'path', path: a.image_path };
  if (a.attachment_id !== undefined)
    return { kind: 'attachment', attachmentId: a.attachment_id };
  return { kind: 'svg', svg: a.svg as string };
}

export const insertImage = defineCommand({
  name: 'insert_image',
  title: 'Insert a picture',
  description:
    'Put a picture in its own paragraph, before or after the paragraph holding anchor_text, or at the end ' +
    'of the document. ' +
    ONE_SOURCE +
    'It comes in at its own size, shrunk to the text width when it is wider; pass width to set it. ' +
    'To replace a picture that is already there, use replace_image instead — it keeps the layout around it.',
  input: {
    documentId,
    ...imageSourceShape,
    position: z.enum(['before', 'after', 'document_end']),
    anchor_text: z
      .string()
      .optional()
      .describe(
        'Required for before/after: exact text inside the anchor paragraph.',
      ),
    width: z
      .number()
      .positive()
      .optional()
      .describe('Display width in CSS px; the height follows the picture.'),
    alt: z.string().optional().describe('Alt text (Word’s Description).'),
    occurrence,
    expectedVersion,
  },
  effect: 'edit',
  requires: 'images',
  targets: (a) => [a.documentId],
  run: (provider, a) =>
    withSession(provider, a.documentId, async (s) => {
      const source = imageSource(a);
      if (typeof source === 'string') return errorText(source);
      if (a.position !== 'document_end' && !a.anchor_text)
        return errorText(
          'anchor_text is required when position is before/after.',
        );
      const anchor =
        a.position === 'document_end'
          ? ({ position: 'document_end' } as const)
          : ({
              position: a.position,
              text: a.anchor_text as string,
              occurrence: a.occurrence,
            } as const);
      return json(
        await s.insertImage(
          source,
          anchor,
          { width: a.width, alt: a.alt },
          { expectedVersion: a.expectedVersion },
        ),
      );
    }),
});

export const replaceImage = defineCommand({
  name: 'replace_image',
  title: 'Replace a picture',
  description:
    'Swap one picture for another, keeping where it sits: its anchor, its text wrap, and its width unless you ' +
    'pass one (the height follows the new picture, so nothing is squashed). This is how you redraw a diagram ' +
    'without disturbing the page. ' +
    ONE_SOURCE +
    'Read get_document first — blocks list their images and (block_index, image_index) addresses one. ' +
    'Mind what you are replacing: kind "drawing" is art the file describes shape by shape and kind "equation" ' +
    'is a real equation; turning either into a flat picture cannot be undone from the file, so say so first. ' +
    'Block indexes change with every edit — pass expectedVersion from your last read.',
  input: {
    documentId,
    block_index: z
      .number()
      .int()
      .min(0)
      .describe('Index of the block holding the picture (from get_document).'),
    image_index: z
      .number()
      .int()
      .min(0)
      .optional()
      .describe('0-based picture within the block (default 0, the first).'),
    ...imageSourceShape,
    width: z
      .number()
      .positive()
      .optional()
      .describe('Display width in CSS px; default keeps the old one.'),
    alt: z
      .string()
      .optional()
      .describe('Alt text; default keeps the old picture’s.'),
    expectedVersion,
  },
  effect: 'edit',
  requires: 'images',
  targets: (a) => [a.documentId],
  run: (provider, a) =>
    withSession(provider, a.documentId, async (s) => {
      const source = imageSource(a);
      if (typeof source === 'string') return errorText(source);
      return json(
        await s.replaceImage(
          a.block_index,
          a.image_index ?? 0,
          source,
          { width: a.width, alt: a.alt },
          { expectedVersion: a.expectedVersion },
        ),
      );
    }),
});

export const deleteImage = defineCommand({
  name: 'delete_image',
  title: 'Delete a picture',
  description:
    'Remove one picture, addressed as (block_index, image_index) from get_document. Its paragraph stays, so ' +
    'the text around it does not move. A diagram drawn as several separate objects takes one call each — and ' +
    'the indexes shift after every one, so re-read get_document between calls. ' +
    'To put something in a picture’s place, replace_image keeps the layout; this does not.',
  input: {
    documentId,
    block_index: z
      .number()
      .int()
      .min(0)
      .describe('Index of the block holding the picture (from get_document).'),
    image_index: z
      .number()
      .int()
      .min(0)
      .optional()
      .describe('0-based picture within the block (default 0, the first).'),
    expectedVersion,
  },
  effect: 'edit',
  requires: 'images',
  targets: (a) => [a.documentId],
  run: (provider, a) =>
    withSession(provider, a.documentId, async (s) =>
      json(
        await s.deleteImage(a.block_index, a.image_index ?? 0, {
          expectedVersion: a.expectedVersion,
        }),
      ),
    ),
});

export const deleteBlock = defineCommand({
  name: 'delete_block',
  title: 'Delete paragraphs',
  description:
    'Remove whole blocks — paragraphs or headings, with their text and pictures — addressed by block_index from ' +
    'get_document; count removes that many consecutive blocks. replace_text with an empty new_text only empties ' +
    'a paragraph; this takes the paragraph away. A paragraph inside a table cell can go only while the cell keeps ' +
    'another: rows, columns and whole tables are removed with edit_table. Block indexes shift after this call — ' +
    're-read before the next one.',
  input: {
    documentId,
    block_index: z
      .number()
      .int()
      .min(0)
      .describe('Index of the first block to remove (from get_document).'),
    count: z
      .number()
      .int()
      .min(1)
      .optional()
      .describe('How many consecutive blocks, default 1.'),
    expectedVersion,
  },
  effect: 'edit',
  targets: (a) => [a.documentId],
  run: (
    provider,
    { documentId: id, block_index, count, expectedVersion: ver },
  ) =>
    withSession(provider, id, async (s) =>
      json(
        await s.deleteBlocks(block_index, count ?? 1, {
          expectedVersion: ver,
        }),
      ),
    ),
});

export const pageSetup = defineCommand({
  name: 'page_setup',
  title: 'Page setup',
  description:
    'Orientation, paper size, margins and text columns — for one section (1-based) or, without section, for ' +
    'every section — and section breaks. get_document marks each block with its section once there is more ' +
    "than one; check_document reports each section's page. To turn a few pages landscape: a section break " +
    'before them and one after (section_break_after: the block index, one call each), then page_setup with ' +
    'that section and orientation. A break goes in or out on its own call, since the sections are numbered ' +
    'anew after it. margins is "normal" | "narrow" | "moderate" | "wide" or { top, right, bottom, left } lengths.',
  input: {
    documentId,
    section: z
      .number()
      .int()
      .min(1)
      .optional()
      .describe('The section to change, 1-based; omit for all of them.'),
    orientation: z.enum(['portrait', 'landscape']).optional(),
    paper: z
      .enum(['A3', 'A4', 'A5', 'Letter', 'Legal', 'Executive'])
      .optional(),
    margins: z
      .union([
        z.enum(['normal', 'narrow', 'moderate', 'wide']),
        z
          .object({
            top: lengthArg.optional(),
            right: lengthArg.optional(),
            bottom: lengthArg.optional(),
            left: lengthArg.optional(),
          })
          .strict(),
      ])
      .optional()
      .describe(
        'A preset, or lengths per side (a side left out keeps its value).',
      ),
    columns: z.number().int().min(1).max(3).optional(),
    section_break_after: z
      .number()
      .int()
      .min(0)
      .optional()
      .describe(
        'Start a new section after this block (index from get_document, a paragraph outside tables).',
      ),
    new_page: z
      .boolean()
      .optional()
      .describe(
        'With section_break_after: the new section starts on a new page (default) or, false, continues on the same one.',
      ),
    remove_section_break: z
      .number()
      .int()
      .min(1)
      .optional()
      .describe(
        'Remove the break that ends this section (1-based), joining it with the next.',
      ),
    expectedVersion,
  },
  effect: 'edit',
  requires: 'pageSetup',
  targets: (a) => [a.documentId],
  run: (provider, a) =>
    withSession(provider, a.documentId, async (s) => {
      if (
        a.orientation === undefined &&
        a.paper === undefined &&
        a.margins === undefined &&
        a.columns === undefined &&
        a.section_break_after === undefined &&
        a.remove_section_break === undefined
      ) {
        return errorText(
          'Pass at least one change: orientation, paper, margins, columns, section_break_after or remove_section_break.',
        );
      }
      return json(
        await s.pageSetup(
          {
            ...(a.section !== undefined ? { section: a.section } : {}),
            ...(a.orientation ? { orientation: a.orientation } : {}),
            ...(a.paper ? { paper: a.paper } : {}),
            ...(a.margins !== undefined ? { margins: a.margins } : {}),
            ...(a.columns !== undefined ? { columns: a.columns } : {}),
            ...(a.section_break_after !== undefined
              ? {
                  sectionBreakAfter: {
                    blockIndex: a.section_break_after,
                    newPage: a.new_page ?? true,
                  },
                }
              : {}),
            ...(a.remove_section_break !== undefined
              ? { removeSectionBreak: a.remove_section_break }
              : {}),
          },
          { expectedVersion: a.expectedVersion },
        ),
      );
    }),
});

export const editHeaderFooter = defineCommand({
  name: 'edit_header_footer',
  title: 'Edit a header or footer',
  description:
    'Rewrite a header or footer — content is text or blocks, like insert_content, where { field: "page" } is ' +
    'the page number and { field: "pages" } the page count — or replace text inside it (old_text, new_text; ' +
    'formatting kept, same matching rules as replace_text). For one section (1-based) or, without section, ' +
    'every section. get_document lists them under chrome ({page} marks a page number). variant "first" or ' +
    '"even" edits a first-page or even-page one the document already has. No lists inside.',
  input: {
    documentId,
    part: z.enum(['header', 'footer']),
    variant: z.enum(['default', 'first', 'even']).optional(),
    section: z
      .number()
      .int()
      .min(1)
      .optional()
      .describe('The section, 1-based; omit for every section.'),
    content: contentSchema
      .optional()
      .describe('The new header / footer, replacing it whole ("" empties it).'),
    old_text: z
      .string()
      .min(1)
      .optional()
      .describe('Instead of content: exact text inside it to replace.'),
    new_text: z.string().optional().describe('The replacement for old_text.'),
    occurrence,
    expectedVersion,
  },
  effect: 'edit',
  requires: 'headerFooter',
  targets: (a) => [a.documentId],
  run: (provider, a) =>
    withSession(provider, a.documentId, async (s) => {
      if (a.old_text !== undefined && a.new_text === undefined)
        return errorText('old_text needs new_text ("" deletes it).');
      return json(
        await s.editChrome(
          {
            part: a.part,
            ...(a.variant ? { variant: a.variant } : {}),
            ...(a.section !== undefined ? { section: a.section } : {}),
            ...(a.content !== undefined ? { content: a.content } : {}),
            ...(a.old_text !== undefined
              ? {
                  replace: {
                    oldText: a.old_text,
                    newText: a.new_text ?? '',
                    ...(a.occurrence !== undefined
                      ? { occurrence: a.occurrence }
                      : {}),
                  },
                }
              : {}),
          },
          { expectedVersion: a.expectedVersion },
        ),
      );
    }),
});

export const listStyles = defineCommand({
  name: 'list_styles',
  title: "List the document's paragraph styles",
  description:
    "The paragraph styles this document defines — Word's built-ins and its own (custom: true) — with how many " +
    "paragraphs use each. Use a name with apply_formatting style or a block's style: the paragraph takes the " +
    "style's look, and Word shows it under that style. Prefer the document's own styles over formatting by hand.",
  input: { documentId },
  effect: 'read',
  requires: 'styles',
  targets: (a) => [a.documentId],
  run: (provider, { documentId: id }) =>
    withSession(provider, id, async (s) =>
      json({ styles: await s.listStyles() }),
    ),
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
 *  needs the `selection` capability, the three picture commands need
 *  `images`, `page_setup` needs `pageSetup`, `edit_header_footer`
 *  `headerFooter` and `list_styles` `styles`; hosts without one leave them
 *  out. */
export const documentCommands: readonly AgentCommand<
  z.ZodRawShape,
  SessionProvider
>[] = [
  getDocument,
  findText,
  listStyles,
  replaceText,
  insertContent,
  deleteBlock,
  applyFormatting,
  editTable,
  updateImage,
  insertImage,
  replaceImage,
  deleteImage,
  pageSetup,
  editHeaderFooter,
  saveDocument,
  getSelection,
];
