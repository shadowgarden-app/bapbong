/**
 * Structured content — what an agent hands to insert_content / create_document
 * when plain lines are not enough: headings, tabs with leaders, tables.
 *
 * The vocabulary is docx-js's (heading, bold, colspan, leader) so a model
 * that already writes docx-js needs nothing new; the axis names are this
 * package's (align, as in apply_formatting). The builder is pure — blocks in,
 * ProseMirror nodes out — and lives in the browser-safe tier: the desktop
 * WebView runs it against its live editor, the headless session against an
 * EditorState, the shell against a fresh document. Validation of the SHAPE is
 * zod's job (./blocks-schema, host tier); this file checks what a schema
 * cannot — ragged rows, a widths list that does not fit — and reports it as
 * a {@link ContentError} the tool layer turns into an error result.
 */
import type { Node as PMNode, Schema, Mark } from 'prosemirror-model';
import type {
  ResolvedTableStyle,
  TableLook,
} from '@shadow-garden/bapbong-headless';
import { ContentError } from './contract.js';

export type Align = 'left' | 'center' | 'right' | 'justify';

/** A run of text with character marks, or a tab. */
export type Inline =
  | string
  | { text: string; bold?: boolean; italic?: boolean; underline?: boolean }
  | { tab: true };

/** A tab stop of the paragraph. `at` is a length: a number is centimetres,
 *  a string may be "100%" (of the text width), "3cm", "1in", "40px". */
export interface TabStop {
  at: number | string;
  align?: 'left' | 'right' | 'center';
  leader?: 'dot' | 'underscore' | 'hyphen';
}

export interface ParagraphBlock {
  paragraph: string | Inline[];
  /** Word "Heading N", 1–6. */
  heading?: number;
  /** Named paragraph style without an outline level. */
  style?: 'Title' | 'Subtitle';
  align?: Align;
  tabs?: TabStop[];
  pageBreakBefore?: boolean;
  /** Marks for the whole paragraph (inline marks add to them). */
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
}

export type Cell =
  | string
  | Inline[]
  | {
      text: string | Inline[];
      /** Spans this many grid columns to the right. */
      colspan?: number;
      align?: Align;
      /** Fill colour, "#RRGGBB". */
      shading?: string;
      vAlign?: 'center' | 'bottom';
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
    };

export interface TableBlock {
  /** Rows of cells; every row must cover the same number of grid columns. */
  table: Cell[][];
  /** One length per grid column (cm as a number, or "%"/"cm"/"in"/"px"
   *  strings); absent = equal columns across the text width. */
  widths?: (number | string)[];
  /** grid (default): every line; outer: the frame only; none: invisible. */
  borders?: 'grid' | 'none' | 'outer';
  /** First row is a header: bold, repeated on every page. */
  header?: boolean;
  /** Table alignment on the page. */
  align?: 'center' | 'right';
}

export type Block = string | ParagraphBlock | TableBlock;
/** What insert_content / create_document accept. */
export type Content = string | Block[];

/** One edit_table call. Applied in this order: deleteRows, insertRows,
 *  merge, widths, then borders / header / align — so indexes in one call
 *  refer to the table as get_document showed it, except that insertRows.at
 *  counts rows AFTER the deletions. */
export interface TableEdit {
  /** Rows (same cell grammar as a table block) inserted before row `at`;
   *  omit `at` to append. */
  insertRows?: { at?: number; rows: Cell[][] };
  /** 0-based row indexes to remove (at least one row must remain). */
  deleteRows?: number[];
  /** Merge cells `from`..`to` (0-based, inclusive) of one row into one. */
  merge?: { row: number; from: number; to: number };
  /** One length per grid column. */
  widths?: (number | string)[];
  borders?: 'grid' | 'none' | 'outer';
  /** Row 0 is a header row, repeated on every page. */
  header?: boolean;
  align?: 'left' | 'center' | 'right';
}

/** The "Table Grid" style a new table is born with, when the host can
 *  resolve it (the docx catalog). Without it a direct 1px grid is used. */
export interface TableStyleSource {
  styleId: string;
  look?: TableLook;
  style?: ResolvedTableStyle;
}

export interface BuildOptions {
  /** Width of the text area in CSS px — what "100%" and equal columns mean. */
  contentWidth: number;
  tableStyle?: TableStyleSource;
}

const PX_PER_CM = 96 / 2.54;
const DEFAULT_LOOK: TableLook = {
  firstRow: true,
  lastRow: false,
  firstCol: true,
  lastCol: false,
  hBand: true,
  vBand: false,
};
const GRID_SIDE = { width: 1, style: 'solid', color: '#000000' } as const;

/** Length → CSS px. */
export function lengthToPx(v: number | string, full: number): number {
  if (typeof v === 'number') return v * PX_PER_CM;
  const m = /^\s*(-?\d+(?:\.\d+)?)\s*(%|cm|mm|in|px|pt)?\s*$/.exec(v);
  if (!m)
    throw new ContentError(
      `Not a length: ${JSON.stringify(v)} — use a number (cm) or "50%", "3cm", "1in", "40px".`,
    );
  const n = Number(m[1]);
  switch (m[2]) {
    case '%':
      return (n / 100) * full;
    case 'mm':
      return (n / 10) * PX_PER_CM;
    case 'in':
      return n * 96;
    case 'px':
      return n;
    case 'pt':
      return (n * 96) / 72;
    default:
      return n * PX_PER_CM;
  }
}

/** Build the ProseMirror nodes for `content`. A plain string becomes one
 *  paragraph per line (the original behaviour); blocks become what they say. */
export function contentToNodes(
  content: Content,
  schema: Schema,
  opts: BuildOptions,
): PMNode[] {
  if (typeof content === 'string') return linesToParagraphs(content, schema);
  const out: PMNode[] = [];
  content.forEach((block, i) => {
    try {
      out.push(...blockToNodes(block, schema, opts));
    } catch (err) {
      if (err instanceof ContentError) {
        throw new ContentError(`Block ${i}: ${err.message}`);
      }
      throw err;
    }
  });
  return out;
}

function blockToNodes(
  block: Block,
  schema: Schema,
  opts: BuildOptions,
): PMNode[] {
  if (typeof block === 'string') return linesToParagraphs(block, schema);
  if ('table' in block) return [tableNode(block, schema, opts)];
  if ('paragraph' in block) return [paragraphNode(block, schema, opts)];
  throw new ContentError(
    'A block is a string, { paragraph: … } or { table: … }.',
  );
}

function linesToParagraphs(text: string, schema: Schema): PMNode[] {
  return text
    .split('\n')
    .map((line) =>
      schema.node(
        'paragraph',
        null,
        line.length > 0 ? [schema.text(line)] : [],
      ),
    );
}

type MarkFlags = { bold?: boolean; italic?: boolean; underline?: boolean };

function marksFor(schema: Schema, ...flags: MarkFlags[]): Mark[] {
  const on = (k: keyof MarkFlags) => flags.some((f) => f[k]);
  const out: Mark[] = [];
  const add = (name: string) => {
    const type = schema.marks[name];
    if (type) out.push(type.create());
  };
  if (on('bold')) add('strong');
  if (on('italic')) add('em');
  if (on('underline')) add('underline');
  return out;
}

/** Inline content → text nodes (a `{ tab }` is a tab character, which the
 *  layout resolves against the paragraph's tab stops). */
function inlineNodes(
  text: string | Inline[],
  schema: Schema,
  base: MarkFlags,
): PMNode[] {
  const runs = typeof text === 'string' ? [text] : text;
  const out: PMNode[] = [];
  for (const run of runs) {
    if (typeof run === 'string') {
      if (run.includes('\n')) {
        throw new ContentError(
          'No "\\n" inside a paragraph — make separate blocks instead.',
        );
      }
      if (run.length > 0) out.push(schema.text(run, marksFor(schema, base)));
    } else if ('tab' in run) {
      out.push(schema.text('\t', marksFor(schema, base)));
    } else {
      if (run.text.includes('\n')) {
        throw new ContentError(
          'No "\\n" inside a paragraph — make separate blocks instead.',
        );
      }
      if (run.text.length > 0)
        out.push(schema.text(run.text, marksFor(schema, base, run)));
    }
  }
  return out;
}

function paragraphAttrs(
  p: {
    heading?: number;
    style?: string;
    align?: Align;
    tabs?: TabStop[];
    pageBreakBefore?: boolean;
  },
  opts: BuildOptions,
): Record<string, unknown> {
  const attrs: Record<string, unknown> = {};
  if (p.heading) attrs['heading'] = p.heading;
  else if (p.style) attrs['styleId'] = p.style;
  if (p.align) attrs['align'] = p.align;
  if (p.pageBreakBefore) attrs['pageBreakBefore'] = true;
  if (p.tabs?.length) {
    attrs['tabs'] = p.tabs.map((t) => ({
      pos: Math.round(lengthToPx(t.at, opts.contentWidth)),
      val: t.align ?? 'left',
      ...(t.leader ? { leader: t.leader } : {}),
    }));
  }
  return attrs;
}

function paragraphNode(
  p: ParagraphBlock,
  schema: Schema,
  opts: BuildOptions,
): PMNode {
  return schema.node(
    'paragraph',
    paragraphAttrs(p, opts),
    inlineNodes(p.paragraph, schema, p),
  );
}

type CellSpec = Exclude<Cell, string | Inline[]>;

function normalizeCell(cell: Cell): CellSpec {
  if (typeof cell === 'string' || Array.isArray(cell)) return { text: cell };
  return cell;
}

function cellSpan(c: CellSpec): number {
  return Math.max(1, Math.floor(c.colspan ?? 1));
}

/** How many grid columns a row of cells covers. */
export function rowSpan(cells: Cell[]): number {
  return cells.map(normalizeCell).reduce((n, c) => n + cellSpan(c), 0);
}

/** Lengths → one px width per grid column; equal columns when absent. */
export function columnWidths(
  widths: (number | string)[] | undefined,
  cols: number,
  contentWidth: number,
): number[] {
  if (widths) {
    if (widths.length !== cols) {
      throw new ContentError(
        `widths lists ${widths.length} column(s) but the table has ${cols}.`,
      );
    }
    return widths.map((w) => Math.round(lengthToPx(w, contentWidth)));
  }
  return Array.from({ length: cols }, () => Math.round(contentWidth / cols));
}

/** The grid columns of an existing table: how many, and their px widths as
 *  the first row's cells declare them (equal columns when they do not). */
export function tableGrid(
  table: PMNode,
  contentWidth: number,
): { cols: number; widths: number[] } {
  const first = table.firstChild;
  if (!first) return { cols: 0, widths: [] };
  const widths: number[] = [];
  let cols = 0;
  first.forEach((cell) => {
    const span = Math.max(1, Number(cell.attrs['colspan']) || 1);
    const cw = cell.attrs['colwidth'] as number[] | null;
    cols += span;
    if (cw && cw.length === span) widths.push(...cw);
    else for (let i = 0; i < span; i++) widths.push(NaN);
  });
  if (widths.some((w) => !Number.isFinite(w) || w <= 0)) {
    return {
      cols,
      widths: Array.from({ length: cols }, () =>
        Math.round(contentWidth / cols),
      ),
    };
  }
  return { cols, widths };
}

/** One table row from cells, laid over `widths` (one per grid column). */
export function rowNode(
  cells: Cell[],
  widths: number[],
  schema: Schema,
  header = false,
): PMNode {
  const { table_row, table_cell } = schema.nodes;
  if (!table_row || !table_cell)
    throw new ContentError('This document schema has no tables.');
  const specs = cells.map(normalizeCell);
  const covered = specs.reduce((n, c) => n + cellSpan(c), 0);
  if (covered !== widths.length) {
    throw new ContentError(
      `A row covers ${covered} column(s) but the table has ${widths.length} — every row must cover the same columns (use colspan to merge).`,
    );
  }
  let col = 0;
  const nodes = specs.map((c) => {
    const n = cellSpan(c);
    const attrs: Record<string, unknown> = {
      colspan: n,
      colwidth: widths.slice(col, col + n),
    };
    col += n;
    if (c.shading) attrs['background'] = c.shading;
    if (c.vAlign) attrs['vAlign'] = c.vAlign;
    const pAttrs = c.align ? { align: c.align } : null;
    const marks: MarkFlags = { ...c, bold: c.bold || header };
    return table_cell.create(
      attrs,
      schema.node('paragraph', pAttrs, inlineNodes(c.text, schema, marks)),
    );
  });
  return table_row.create(header ? { header: true } : null, nodes);
}

/** The table attrs that draw its lines: a style for `grid` when the host has
 *  one (else direct borders), the frame only for `outer`, nothing for `none`. */
export function tableChrome(
  borders: 'grid' | 'none' | 'outer',
  tableStyle: TableStyleSource | undefined,
  styleable: boolean,
): {
  styleId: string | null;
  look: TableLook | null;
  borders: Record<string, unknown> | null;
} {
  if (borders === 'grid') {
    if (tableStyle && styleable) {
      return {
        styleId: tableStyle.styleId,
        look: { ...(tableStyle.look ?? DEFAULT_LOOK) },
        borders: null,
      };
    }
    return {
      styleId: null,
      look: null,
      borders: {
        top: GRID_SIDE,
        bottom: GRID_SIDE,
        left: GRID_SIDE,
        right: GRID_SIDE,
        insideH: GRID_SIDE,
        insideV: GRID_SIDE,
      },
    };
  }
  if (borders === 'outer') {
    return {
      styleId: null,
      look: null,
      borders: {
        top: GRID_SIDE,
        bottom: GRID_SIDE,
        left: GRID_SIDE,
        right: GRID_SIDE,
        insideH: null,
        insideV: null,
      },
    };
  }
  return { styleId: null, look: null, borders: null };
}

function tableNode(t: TableBlock, schema: Schema, opts: BuildOptions): PMNode {
  const { table } = schema.nodes;
  if (!table) throw new ContentError('This document schema has no tables.');
  if (t.table.length === 0)
    throw new ContentError('A table needs at least one row.');
  const cols = rowSpan(t.table[0]);
  t.table.forEach((r, i) => {
    const n = rowSpan(r);
    if (n !== cols) {
      throw new ContentError(
        `Row ${i} covers ${n} column(s), row 0 covers ${cols} — every row must cover the same columns (use colspan to merge).`,
      );
    }
  });
  const widths = columnWidths(t.widths, cols, opts.contentWidth);
  const rows = t.table.map((r, ri) =>
    rowNode(r, widths, schema, !!t.header && ri === 0),
  );
  const chrome = tableChrome(
    t.borders ?? 'grid',
    opts.tableStyle,
    !!table.spec.attrs?.['styleId'],
  );
  const attrs: Record<string, unknown> = {};
  if (chrome.styleId) attrs['styleId'] = chrome.styleId;
  if (chrome.look) attrs['look'] = chrome.look;
  if (chrome.borders) attrs['borders'] = chrome.borders;
  if (t.align) attrs['align'] = t.align;
  return table.create(attrs, rows);
}

/** Table style ids the nodes reference — the caller injects their
 *  definitions into doc.attrs.tableStyles when the sheet lacks them. */
export function referencedTableStyles(nodes: PMNode[]): string[] {
  const ids = new Set<string>();
  for (const n of nodes) {
    n.descendants((d) => {
      const id =
        d.type.name === 'table' ? (d.attrs['styleId'] as string | null) : null;
      if (id) ids.add(id);
      return true;
    });
    if (n.type.name === 'table' && n.attrs['styleId'])
      ids.add(String(n.attrs['styleId']));
  }
  return [...ids];
}
