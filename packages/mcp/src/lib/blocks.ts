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

/** Character formatting a run of text (or a whole paragraph or cell) can
 *  carry. Flags add up from paragraph to run; a value set on a run (colour,
 *  font, size) wins over the paragraph's. */
export interface CharFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  /** Text colour, "#RRGGBB". */
  color?: string;
  /** Highlight (the colour behind the text), "#RRGGBB". */
  highlight?: string;
  /** Font family name, as Word shows it ("Arial", "Times New Roman"). */
  font?: string;
  /** Font size in points. */
  size?: number;
}

/** A run of text with character formatting — and, optionally, a hyperlink
 *  (a web address, "mailto:…", or "#bookmark") — or a tab. */
export type Inline =
  | string
  | ({ text: string; link?: string } & CharFormat)
  | { tab: true };

/**
 * A link target an agent may write: http(s), mailto, tel, or a bookmark in
 * the document ("#name"). A bare address gets https:// the way the editor's
 * link panel adds it. Anything else — javascript:, file:, data: — is refused:
 * the user will click it.
 */
export function linkTarget(href: string): string {
  const t = href.trim();
  if (!t) throw new ContentError('A link needs an address.');
  if (t.startsWith('#')) return t;
  const scheme = /^([a-z][a-z0-9+.-]*):/i.exec(t)?.[1]?.toLowerCase();
  if (!scheme) return `https://${t}`;
  if (['http', 'https', 'mailto', 'tel'].includes(scheme)) return t;
  throw new ContentError(
    `A link goes to a web address, "mailto:", "tel:" or "#bookmark" — not ${JSON.stringify(`${scheme}:`)}.`,
  );
}

/** A tab stop of the paragraph. `at` is a length: a number is centimetres,
 *  a string may be "100%" (of the text width), "3cm", "1in", "40px". */
export interface TabStop {
  at: number | string;
  align?: 'left' | 'right' | 'center';
  leader?: 'dot' | 'underscore' | 'hyphen';
}

/** Character formatting on a paragraph block applies to all its text
 *  (inline formatting adds to it). */
export interface ParagraphBlock extends CharFormat, ParagraphLayout {
  paragraph: string | Inline[];
  /** Word "Heading N", 1–6. */
  heading?: number;
  /** Named paragraph style without an outline level. */
  style?: 'Title' | 'Subtitle';
  align?: Align;
  tabs?: TabStop[];
  pageBreakBefore?: boolean;
  /** A list item: a real Word list, never a typed "•" or "1.". Consecutive
   *  items of one kind are one list — numbers count on from the item above. */
  list?: ListKind;
  /** Nesting level of a list item, 1 (top) to {@link LIST_LEVELS}. */
  level?: number;
}

/** Paragraph spacing and indents as an agent writes them (see
 *  {@link spacingAttr} / {@link indentAttr}). */
export interface ParagraphLayout {
  /** Space above / below, points. */
  spaceBefore?: number;
  spaceAfter?: number;
  /** A multiple of single (1, 1.15, 1.5, 2) or an exact / minimum height in
   *  points. */
  lineSpacing?: number | { exact: number } | { atLeast: number };
  /** Lengths: cm as a number, or "1cm", "0.5in", "12pt". 0 removes one. */
  indent?: {
    left?: number | string;
    right?: number | string;
    firstLine?: number | string;
    hanging?: number | string;
  };
}

const PX_PER_PT = 96 / 72;

/** The paragraph's `spacing` attr (px; line a multiple under lineRule
 *  auto) with `p`'s changes over `old`, or undefined when `p` changes none.
 *  A side set here drops its "auto" flag: it is direct formatting now. */
export function spacingAttr(
  old: Record<string, unknown> | null | undefined,
  p: ParagraphLayout,
): Record<string, unknown> | null | undefined {
  if (
    p.spaceBefore === undefined &&
    p.spaceAfter === undefined &&
    p.lineSpacing === undefined
  )
    return undefined;
  const next: Record<string, unknown> = { ...(old ?? {}) };
  const pt = (v: number, what: string) => {
    if (!Number.isFinite(v) || v < 0 || v > 1584)
      throw new ContentError(`${what} is points, 0 to 1584 — not ${v}.`);
    return Math.round(v * PX_PER_PT);
  };
  if (p.spaceBefore !== undefined) {
    next['before'] = pt(p.spaceBefore, 'space_before');
    delete next['beforeAuto'];
  }
  if (p.spaceAfter !== undefined) {
    next['after'] = pt(p.spaceAfter, 'space_after');
    delete next['afterAuto'];
  }
  const ls = p.lineSpacing;
  if (typeof ls === 'number') {
    if (!Number.isFinite(ls) || ls < 0.5 || ls > 10)
      throw new ContentError(
        `line_spacing is a multiple of single line spacing, 0.5 to 10 — not ${ls}. For a height in points, pass "18pt".`,
      );
    next['line'] = ls;
    next['lineRule'] = 'auto';
  } else if (ls && 'exact' in ls) {
    next['line'] = pt(ls.exact, 'line_spacing');
    next['lineRule'] = 'exact';
  } else if (ls && 'atLeast' in ls) {
    next['line'] = pt(ls.atLeast, 'line_spacing');
    next['lineRule'] = 'atLeast';
  }
  return Object.keys(next).length > 0 ? next : null;
}

/** The paragraph's `indent` attr (px) with `change` over `old`, or
 *  undefined when it changes nothing. 0 removes a side; firstLine and
 *  hanging exclude each other. */
export function indentAttr(
  old: Record<string, unknown> | null | undefined,
  change: ParagraphLayout['indent'],
  contentWidth: number,
): Record<string, unknown> | null | undefined {
  if (!change) return undefined;
  const next: Record<string, unknown> = { ...(old ?? {}) };
  let touched = false;
  for (const side of ['left', 'right', 'firstLine', 'hanging'] as const) {
    const v = change[side];
    if (v === undefined) continue;
    touched = true;
    const px = Math.round(lengthToPx(v, contentWidth));
    if (side === 'firstLine' || side === 'hanging') {
      delete next['firstLine'];
      delete next['hanging'];
      if (px < 0)
        throw new ContentError(
          `${side === 'firstLine' ? 'first_line' : 'hanging'} is a length of 0 or more — for the other direction use ${side === 'firstLine' ? 'hanging' : 'first_line'}.`,
        );
    }
    if (px === 0) delete next[side];
    else next[side] = px;
  }
  if (!touched) return undefined;
  return Object.keys(next).length > 0 ? next : null;
}

/** What an agent calls the two kinds of list. */
export type ListKind = 'bullet' | 'number';

export type Cell =
  | string
  | Inline[]
  | ({
      text: string | Inline[];
      /** Spans this many grid columns to the right. */
      colspan?: number;
      align?: Align;
      /** Fill colour, "#RRGGBB". */
      shading?: string;
      vAlign?: 'center' | 'bottom';
    } & CharFormat);

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

/** One edit_table call. Applied in this order: deleteRows, deleteColumns,
 *  insertColumns, insertRows, merge, widths, then borders / header / align —
 *  so indexes in one call refer to the table as get_document showed it,
 *  except that insertColumns.at and insertRows.at count AFTER the deletions,
 *  and inserted rows cover the columns the table has by then. */
export interface TableEdit {
  /** Remove the whole table. Alone in its call. */
  deleteTable?: boolean;
  /** Empty columns inserted before grid column `at` (omit to append),
   *  `count` of them (default 1). The table keeps its width. */
  insertColumns?: { at?: number; count?: number };
  /** 0-based grid columns to remove (at least one must remain). */
  deleteColumns?: number[];
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
  /** The document's numbering definitions (`doc.attrs.numbering`), so a new
   *  list gets an id nobody uses and joins the right definitions. */
  numbering?: NumberingDefs | null;
  /** The list item right before the insertion point: a first list item of
   *  the same kind joins its list instead of starting a new one. */
  continueList?: { kind: ListKind; numId: string } | null;
}

/** The subset of the model's NumberingDefs this tier reads and writes
 *  (plain data; the model package owns the full type). */
export type NumberingDefs = Record<
  string,
  {
    key: string;
    levels: Record<
      number,
      { numFmt: string; lvlText: string; start?: number } | undefined
    >;
  }
>;

/** How deep a list an agent makes may nest: the editor's own lists define
 *  three levels, and Tab in the editor stops there too. */
export const LIST_LEVELS = 3;

/** Indent per nesting level — what Tab in the editor adds (0.25"). */
export const LIST_LEVEL_INDENT = 24;

/**
 * The definitions a new list is born with — the editor's own defaults
 * (commands `listPresets(…)[0]`: • ◦ ▪ and 1. a. i.), repeated here because
 * this tier carries no runtime dependency; a spec pins them to the presets.
 *
 * Bullets share the editor's `bb-bullet` id (nothing counts). Every numbered
 * list gets an id and a counter of its own, so a second list starts at 1
 * again instead of counting on from the first; the `bb-ordered` prefix is
 * what tells the editor's list buttons it is numbered.
 */
export const LIST_DEFS = {
  bullet: {
    numId: 'bb-bullet',
    levels: {
      0: { numFmt: 'bullet', lvlText: '•', start: 1 },
      1: { numFmt: 'bullet', lvlText: '◦', start: 1 },
      2: { numFmt: 'bullet', lvlText: '▪', start: 1 },
    },
  },
  number: {
    numId: 'bb-ordered',
    levels: {
      0: { numFmt: 'decimal', lvlText: '%1.', start: 1 },
      1: { numFmt: 'lowerLetter', lvlText: '%2.', start: 1 },
      2: { numFmt: 'lowerRoman', lvlText: '%3.', start: 1 },
    },
  },
} as const;

/** Which kind a list paragraph's definition is, at `level` (0-based):
 *  bullet when that level draws a glyph, number otherwise. */
export function listKindOf(
  numId: string,
  level: number,
  defs: NumberingDefs | null | undefined,
): ListKind {
  if (numId.startsWith('bb-bullet')) return 'bullet';
  if (numId.startsWith('bb-ordered')) return 'number';
  const d = defs?.[numId]?.levels;
  const fmt = d?.[level]?.numFmt ?? d?.[0]?.numFmt;
  return fmt === 'bullet' ? 'bullet' : 'number';
}

/** Add a new list of `kind` to `defs` (mutated) and return its id. */
export function mintList(kind: ListKind, defs: NumberingDefs): string {
  const base = LIST_DEFS[kind];
  let numId: string = base.numId;
  if (kind === 'number') {
    for (let n = 2; defs[numId]; n++) numId = `${base.numId}-n${n}`;
  }
  if (!defs[numId]) {
    defs[numId] = {
      key: numId,
      levels: { ...base.levels },
    };
  }
  return numId;
}

/** Content built for a document: the nodes, and the numbering definitions
 *  the document must carry afterwards (null when no list was added). */
export interface BuiltContent {
  nodes: PMNode[];
  numbering: NumberingDefs | null;
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
 *  paragraph per line (the original behaviour); blocks become what they say.
 *  Content with list items needs {@link buildContent}, which also hands back
 *  the numbering definitions the document must carry. */
export function contentToNodes(
  content: Content,
  schema: Schema,
  opts: BuildOptions,
): PMNode[] {
  return buildContent(content, schema, opts).nodes;
}

/** {@link contentToNodes}, plus the numbering definitions its lists need:
 *  set `numbering` as the document's `numbering` attr when it is not null. */
export function buildContent(
  content: Content,
  schema: Schema,
  opts: BuildOptions,
): BuiltContent {
  if (typeof content === 'string')
    return { nodes: linesToParagraphs(content, schema), numbering: null };
  const defs: NumberingDefs = { ...(opts.numbering ?? {}) };
  const before = Object.keys(defs).length;
  const lists: ListRun = { current: opts.continueList ?? null, defs };
  const out: PMNode[] = [];
  content.forEach((block, i) => {
    try {
      out.push(...blockToNodes(block, schema, opts, lists));
    } catch (err) {
      if (err instanceof ContentError) {
        throw new ContentError(`Block ${i}: ${err.message}`);
      }
      throw err;
    }
  });
  return {
    nodes: out,
    numbering: Object.keys(defs).length > before ? defs : null,
  };
}

/** The list being built while blocks are read: consecutive items of one
 *  kind are one list; anything else between them ends it. */
interface ListRun {
  current: { kind: ListKind; numId: string } | null;
  defs: NumberingDefs;
}

function blockToNodes(
  block: Block,
  schema: Schema,
  opts: BuildOptions,
  lists: ListRun,
): PMNode[] {
  if (typeof block === 'string') {
    lists.current = null;
    return linesToParagraphs(block, schema);
  }
  if ('table' in block) {
    lists.current = null;
    return [tableNode(block, schema, opts)];
  }
  if ('paragraph' in block) {
    return [paragraphNode(block, schema, opts, listAttrs(block, lists))];
  }
  throw new ContentError(
    'A block is a string, { paragraph: … } or { table: … }.',
  );
}

/** The `list` + `indent` attrs of a paragraph block (none when it is not a
 *  list item), joining or starting the list it belongs to. */
function listAttrs(p: ParagraphBlock, lists: ListRun): Record<string, unknown> {
  if (!p.list) {
    if (p.level !== undefined)
      throw new ContentError('level is for list items — add list too.');
    lists.current = null;
    return {};
  }
  const level = listLevel(p.level);
  if (lists.current?.kind !== p.list) {
    lists.current = { kind: p.list, numId: mintList(p.list, lists.defs) };
  }
  return {
    list: { numId: lists.current.numId, level },
    ...(level > 0 ? { indent: { left: level * LIST_LEVEL_INDENT } } : {}),
  };
}

/** A 1-based level from an agent → the model's 0-based one. */
export function listLevel(level: number | undefined): number {
  const l = level ?? 1;
  if (!Number.isInteger(l) || l < 1 || l > LIST_LEVELS) {
    throw new ContentError(
      `List level ${level} is out of range — lists go 1 to ${LIST_LEVELS} levels deep.`,
    );
  }
  return l - 1;
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

type MarkFlags = CharFormat;

/** The marks for a run: flags from any level add up; for a value (colour,
 *  font, size) the last level that sets one — the run — wins. */
function marksFor(schema: Schema, ...flags: MarkFlags[]): Mark[] {
  const on = (k: keyof MarkFlags) => flags.some((f) => f[k]);
  const last = <K extends 'color' | 'highlight' | 'font' | 'size'>(k: K) =>
    flags.reduce<MarkFlags[K]>((v, f) => f[k] ?? v, undefined);
  const out: Mark[] = [];
  const add = (name: string, attrs?: Record<string, unknown>) => {
    const type = schema.marks[name];
    if (type) out.push(type.create(attrs));
  };
  if (on('bold')) add('strong');
  if (on('italic')) add('em');
  if (on('underline')) add('underline');
  if (on('strike')) add('strike');
  if (on('superscript')) add('vertAlign', { value: 'super' });
  else if (on('subscript')) add('vertAlign', { value: 'sub' });
  const color = last('color');
  if (color) add('textColor', { color });
  const highlight = last('highlight');
  if (highlight) add('highlight', { color: highlight });
  const font = last('font');
  if (font) add('fontFamily', { family: font });
  const size = last('size');
  if (size) add('fontSize', { size });
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
      if (run.text.length > 0) {
        const marks = marksFor(schema, base, run);
        const link = schema.marks['link'];
        if (run.link && link)
          marks.push(link.create({ href: linkTarget(run.link) }));
        out.push(schema.text(run.text, marks));
      }
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
  } & ParagraphLayout,
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
  const spacing = spacingAttr(null, p);
  if (spacing) attrs['spacing'] = spacing;
  const indent = indentAttr(null, p.indent, opts.contentWidth);
  if (indent) attrs['indent'] = indent;
  return attrs;
}

function paragraphNode(
  p: ParagraphBlock,
  schema: Schema,
  opts: BuildOptions,
  list: Record<string, unknown> = {},
): PMNode {
  const attrs = paragraphAttrs(p, opts);
  // A list level's indent is the base; an explicit indent on the block wins.
  const indent =
    list['indent'] || attrs['indent']
      ? {
          ...((list['indent'] as object | undefined) ?? {}),
          ...((attrs['indent'] as object | undefined) ?? {}),
        }
      : undefined;
  return schema.node(
    'paragraph',
    { ...attrs, ...list, ...(indent ? { indent } : {}) },
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
