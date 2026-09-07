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
  if (!m) throw new ContentError(`Not a length: ${JSON.stringify(v)} — use a number (cm) or "50%", "3cm", "1in", "40px".`);
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

function blockToNodes(block: Block, schema: Schema, opts: BuildOptions): PMNode[] {
  if (typeof block === 'string') return linesToParagraphs(block, schema);
  if ('table' in block) return [tableNode(block, schema, opts)];
  if ('paragraph' in block) return [paragraphNode(block, schema, opts)];
  throw new ContentError('A block is a string, { paragraph: … } or { table: … }.');
}

function linesToParagraphs(text: string, schema: Schema): PMNode[] {
  return text.split('\n').map((line) =>
    schema.node('paragraph', null, line.length > 0 ? [schema.text(line)] : []),
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
function inlineNodes(text: string | Inline[], schema: Schema, base: MarkFlags): PMNode[] {
  const runs = typeof text === 'string' ? [text] : text;
  const out: PMNode[] = [];
  for (const run of runs) {
    if (typeof run === 'string') {
      if (run.includes('\n')) {
        throw new ContentError('No "\\n" inside a paragraph — make separate blocks instead.');
      }
      if (run.length > 0) out.push(schema.text(run, marksFor(schema, base)));
    } else if ('tab' in run) {
      out.push(schema.text('\t', marksFor(schema, base)));
    } else {
      if (run.text.includes('\n')) {
        throw new ContentError('No "\\n" inside a paragraph — make separate blocks instead.');
      }
      if (run.text.length > 0) out.push(schema.text(run.text, marksFor(schema, base, run)));
    }
  }
  return out;
}

function paragraphAttrs(
  p: { heading?: number; style?: string; align?: Align; tabs?: TabStop[]; pageBreakBefore?: boolean },
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

function paragraphNode(p: ParagraphBlock, schema: Schema, opts: BuildOptions): PMNode {
  return schema.node('paragraph', paragraphAttrs(p, opts), inlineNodes(p.paragraph, schema, p));
}

function normalizeCell(cell: Cell): Exclude<Cell, string | Inline[]> {
  if (typeof cell === 'string' || Array.isArray(cell)) return { text: cell };
  return cell;
}

function tableNode(t: TableBlock, schema: Schema, opts: BuildOptions): PMNode {
  const { table, table_row, table_cell } = schema.nodes;
  if (!table || !table_row || !table_cell) {
    throw new ContentError('This document schema has no tables.');
  }
  if (t.table.length === 0) throw new ContentError('A table needs at least one row.');
  const rows = t.table.map((r) => r.map(normalizeCell));
  const span = (c: ReturnType<typeof normalizeCell>) => Math.max(1, Math.floor(c.colspan ?? 1));
  const cols = rows[0].reduce((n, c) => n + span(c), 0);
  rows.forEach((r, i) => {
    const n = r.reduce((s, c) => s + span(c), 0);
    if (n !== cols) {
      throw new ContentError(
        `Row ${i} covers ${n} column(s), row 0 covers ${cols} — every row must cover the same columns (use colspan to merge).`,
      );
    }
  });
  let widths: number[];
  if (t.widths) {
    if (t.widths.length !== cols) {
      throw new ContentError(`widths lists ${t.widths.length} column(s) but the rows have ${cols}.`);
    }
    widths = t.widths.map((w) => Math.round(lengthToPx(w, opts.contentWidth)));
  } else {
    widths = Array.from({ length: cols }, () => Math.round(opts.contentWidth / cols));
  }

  const rowNodes = rows.map((r, ri) => {
    const isHeader = !!t.header && ri === 0;
    let col = 0;
    const cells = r.map((c) => {
      const n = span(c);
      const attrs: Record<string, unknown> = {
        colspan: n,
        colwidth: widths.slice(col, col + n),
      };
      col += n;
      if (c.shading) attrs['background'] = c.shading;
      if (c.vAlign) attrs['vAlign'] = c.vAlign;
      const pAttrs = c.align ? { align: c.align } : null;
      const marks: MarkFlags = { ...c, bold: c.bold || isHeader };
      const para = schema.node('paragraph', pAttrs, inlineNodes(c.text, schema, marks));
      return table_cell.create(attrs, para);
    });
    return table_row.create(isHeader ? { header: true } : null, cells);
  });

  const attrs: Record<string, unknown> = {};
  const borders = t.borders ?? 'grid';
  if (borders === 'grid') {
    if (opts.tableStyle && table.spec.attrs?.['styleId']) {
      attrs['styleId'] = opts.tableStyle.styleId;
      attrs['look'] = { ...(opts.tableStyle.look ?? DEFAULT_LOOK) };
    } else {
      attrs['borders'] = {
        top: GRID_SIDE,
        bottom: GRID_SIDE,
        left: GRID_SIDE,
        right: GRID_SIDE,
        insideH: GRID_SIDE,
        insideV: GRID_SIDE,
      };
    }
  } else if (borders === 'outer') {
    attrs['borders'] = {
      top: GRID_SIDE,
      bottom: GRID_SIDE,
      left: GRID_SIDE,
      right: GRID_SIDE,
      insideH: null,
      insideV: null,
    };
  }
  if (t.align) attrs['align'] = t.align;
  return table.create(attrs, rowNodes);
}

/** Table style ids the nodes reference — the caller injects their
 *  definitions into doc.attrs.tableStyles when the sheet lacks them. */
export function referencedTableStyles(nodes: PMNode[]): string[] {
  const ids = new Set<string>();
  for (const n of nodes) {
    n.descendants((d) => {
      const id = d.type.name === 'table' ? (d.attrs['styleId'] as string | null) : null;
      if (id) ids.add(id);
      return true;
    });
    if (n.type.name === 'table' && n.attrs['styleId']) ids.add(String(n.attrs['styleId']));
  }
  return [...ids];
}
