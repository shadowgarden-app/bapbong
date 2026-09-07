/**
 * The zod shape of {@link Content} (./blocks) plus the one-paragraph grammar
 * every command that accepts content puts in its description — MCP clients
 * never read a SKILL.md, so the tool text must teach the shape by itself.
 */
import { z } from 'zod';

const marks = {
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
};

const align = z.enum(['left', 'center', 'right', 'justify']);

export const inlineSchema = z.union([
  z.string(),
  z.object({ text: z.string(), ...marks }).strict(),
  z.object({ tab: z.literal(true) }).strict(),
]);

export const lengthSchema = z.union([z.number(), z.string()]);
const length = lengthSchema;

export const tabStopSchema = z
  .object({
    at: length.describe(
      'Position: a number is cm; or "100%" (of the text width), "3cm", "1in".',
    ),
    align: z.enum(['left', 'right', 'center']).optional(),
    leader: z.enum(['dot', 'underscore', 'hyphen']).optional(),
  })
  .strict();

export const paragraphBlockSchema = z
  .object({
    paragraph: z.union([z.string(), z.array(inlineSchema)]),
    heading: z.number().int().min(1).max(6).optional(),
    style: z.enum(['Title', 'Subtitle']).optional(),
    align: align.optional(),
    tabs: z.array(tabStopSchema).optional(),
    pageBreakBefore: z.boolean().optional(),
    ...marks,
  })
  .strict();

export const cellSchema = z.union([
  z.string(),
  z.array(inlineSchema),
  z
    .object({
      text: z.union([z.string(), z.array(inlineSchema)]),
      colspan: z.number().int().min(1).optional(),
      align: align.optional(),
      shading: z
        .string()
        .regex(/^#[0-9a-fA-F]{6}$/)
        .optional(),
      vAlign: z.enum(['center', 'bottom']).optional(),
      ...marks,
    })
    .strict(),
]);

export const tableBlockSchema = z
  .object({
    table: z.array(z.array(cellSchema)).min(1),
    widths: z.array(length).optional(),
    borders: z.enum(['grid', 'none', 'outer']).optional(),
    header: z.boolean().optional(),
    align: z.enum(['center', 'right']).optional(),
  })
  .strict();

export const blockSchema = z.union([
  z.string(),
  paragraphBlockSchema,
  tableBlockSchema,
]);

/** The edit_table input pieces (see TableEdit in ./blocks). */
export const tableEditShape = {
  insert_rows: z
    .object({
      at: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe('Insert before this 0-based row; omit to append.'),
      rows: z
        .array(z.array(cellSchema))
        .min(1)
        .describe('Rows of cells, same grammar as a table block.'),
    })
    .strict()
    .optional(),
  delete_rows: z
    .array(z.number().int().min(0))
    .optional()
    .describe('0-based rows to remove.'),
  merge: z
    .object({
      row: z.number().int().min(0),
      from: z.number().int().min(0),
      to: z.number().int().min(0),
    })
    .strict()
    .optional()
    .describe(
      'Merge cells from..to (0-based, inclusive) of one row into one cell.',
    ),
  widths: z
    .array(length)
    .optional()
    .describe('One length per grid column: cm as a number, or "%".'),
  borders: z.enum(['grid', 'none', 'outer']).optional(),
  header: z
    .boolean()
    .optional()
    .describe('Row 0 is a header row, repeated on every page.'),
  align: z
    .enum(['left', 'center', 'right'])
    .optional()
    .describe('Table alignment on the page.'),
};

/** `content` as commands accept it: plain text, or blocks. */
export const contentSchema = z.union([z.string(), z.array(blockSchema)]);

/** The grammar, for tool descriptions. One paragraph; keep it in step with ./blocks. */
export const CONTENT_GRAMMAR =
  'Plain text (every line one paragraph) or an array of blocks. ' +
  'A block is a string (a paragraph), ' +
  '{ paragraph: text | inlines, heading?: 1-6, style?: "Title"|"Subtitle", align?, ' +
  'tabs?: [{ at: cm | "100%", align?: "right"|"center", leader?: "dot"|"underscore" }], pageBreakBefore? }, ' +
  'or { table: rows, widths?: [cm | "%" per column], borders?: "grid"|"none"|"outer", header?: true, align?: "center" } ' +
  'where a row is a list of cells and a cell is text | inlines | { text, colspan?, align?, bold?, shading?: "#RRGGBB" }. ' +
  'Inlines: text | { text, bold?, italic?, underline? } | { tab: true }. ' +
  'Never draw a table, a column or a dotted line with characters: use a table block ' +
  '(borders "none" for side-by-side text such as signature areas) and a tab with a leader for fill-in lines.';
