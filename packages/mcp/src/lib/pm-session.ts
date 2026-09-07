/**
 * The shared DocumentSession semantics over ANY ProseMirror editor state —
 * anchoring (atom-safe text hits → PM positions), unique-match rules,
 * optimistic locking, mutations as transactions.
 *
 * Hosts provide the state and the dispatch:
 *  - HeadlessSession owns an EditorState (server / tests);
 *  - the desktop WebView wraps its live editor (state + dispatch) so AI edits
 *    ride the normal transaction pipeline (undo, autosave, journal).
 * Keeping ALL semantics here means both hosts behave identically — tested
 * once, headlessly.
 */
import type { EditorState, Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import {
  AnchorError,
  ContentError,
  VersionConflictError,
  type DocBlock,
  type DocSnapshot,
  type DocumentSession,
  type FindMatch,
  type FormatTarget,
  type Formatting,
  type ImageChanges,
  type InsertAnchor,
  type MutationOptions,
  type MutationResult,
  type SessionCapabilities,
} from './contract.js';
import {
  columnWidths,
  contentToNodes,
  lengthToPx,
  referencedTableStyles,
  rowNode,
  tableChrome,
  tableGrid,
  type Content,
  type TableEdit,
  type TableStyleSource,
} from './blocks.js';

/** What a PmDocSession needs from its host. */
export interface PmSessionHost {
  getState(): EditorState;
  /** Dispatch a transaction (live editor) / apply it (headless state). */
  apply(tr: Transaction): void;
  /** Opaque, changes on every doc change — INCLUDING user edits. */
  getVersion(): string;
  meta(): { name?: string; dirty?: boolean };
  save(): Promise<void>;
  /** Current user selection, when the host has one (desktop editor). */
  selection?(): { from: number; to: number } | null;
  /** The style a new table is born with ("Table Grid"), when the host can
   *  resolve it; absent → a direct 1px grid. */
  tableStyle?(): TableStyleSource | undefined;
}

/** A4 with 1in margins — what the layout shows when the doc has no page
 *  attr, so "100%" and equal columns must mean the same here. */
const DEFAULT_CONTENT_WIDTH = 794 - 96 * 2;

interface TextBlock {
  node: PMNode;
  pos: number;
  table?: { index: number; row: number; cell: number };
}

/** A text hit resolved to absolute PM positions. */
interface Hit {
  from: number;
  to: number;
  blockIndex: number;
  context: string;
}

const MARK_BY_FLAG = {
  bold: 'strong',
  italic: 'em',
  underline: 'underline',
  strike: 'strike',
} as const;

export class PmDocSession implements DocumentSession {
  readonly capabilities: SessionCapabilities;

  constructor(private readonly host: PmSessionHost) {
    this.capabilities = { selection: typeof host.selection === 'function' };
  }

  // ── reads ────────────────────────────────────────────────────────────

  async snapshot(): Promise<DocSnapshot> {
    const blocks: DocBlock[] = this.textblocks().map(
      ({ node, table }, index) => {
        const block: DocBlock = {
          index,
          type: blockType(node),
          text: node.textContent,
        };
        if (table) block.table = table;
        const images = blockImages(node).map(({ node: img }, i) => ({
          index: i,
          alt: String(img.attrs['alt'] ?? ''),
          width: Number(img.attrs['width']) || 0,
          height: Number(img.attrs['height']) || 0,
          rotation: Number(img.attrs['rotation']) || 0,
          kind: img.attrs['shape'] ? ('shape' as const) : ('bitmap' as const),
        }));
        if (images.length > 0) block.images = images;
        return block;
      },
    );
    return {
      docVersion: this.host.getVersion(),
      blocks,
      meta: this.host.meta(),
    };
  }

  async find(query: string): Promise<FindMatch[]> {
    return this.hits(query).map((h, i) => ({
      blockIndex: h.blockIndex,
      occurrence: i + 1,
      context: h.context,
    }));
  }

  async getSelection(): Promise<{ text: string; blockIndex: number } | null> {
    const sel = this.host.selection?.();
    if (!sel || sel.from === sel.to) return null;
    const state = this.host.getState();
    const text = state.doc.textBetween(sel.from, sel.to, '\n');
    const blockIndex = this.textblocks().findIndex(
      ({ node, pos }) => sel.from >= pos && sel.from <= pos + node.nodeSize,
    );
    return { text, blockIndex };
  }

  // ── mutations ────────────────────────────────────────────────────────

  async replaceText(
    oldText: string,
    newText: string,
    opts: MutationOptions = {},
  ): Promise<MutationResult> {
    this.checkVersion(opts.expectedVersion);
    const hit = this.uniqueHit(oldText, opts.occurrence);
    this.host.apply(
      this.host.getState().tr.insertText(newText, hit.from, hit.to),
    );
    return {
      docVersion: this.host.getVersion(),
      range: { from: hit.from, to: hit.from + newText.length },
    };
  }

  async insertContent(
    content: Content,
    anchor: InsertAnchor,
    opts: MutationOptions = {},
  ): Promise<MutationResult> {
    this.checkVersion(opts.expectedVersion);
    const state = this.host.getState();
    const { schema } = state;
    const tableStyle = this.host.tableStyle?.();
    const paragraphs = contentToNodes(content, schema, {
      contentWidth: this.contentWidth(),
      tableStyle,
    });

    let insertAt: number;
    if (anchor.position === 'document_end') {
      insertAt = state.doc.content.size;
    } else {
      const hit = this.uniqueHit(anchor.text, anchor.occurrence);
      const block = this.textblocks()[hit.blockIndex];
      insertAt =
        anchor.position === 'before'
          ? block.pos
          : block.pos + block.node.nodeSize;
    }
    const inserted = paragraphs.reduce((size, node) => size + node.nodeSize, 0);
    let tr = state.tr.insert(insertAt, paragraphs);
    // A table born with a style needs its definition in the document's
    // sheet for the layout to paint it (the same move insertTable makes).
    if (tableStyle?.style && schema.nodes['doc'].spec.attrs?.['tableStyles']) {
      const sheet = (state.doc.attrs['tableStyles'] ?? {}) as Record<
        string,
        unknown
      >;
      for (const id of referencedTableStyles(paragraphs)) {
        if (id === tableStyle.styleId && !sheet[id]) {
          tr = tr.setDocAttribute('tableStyles', {
            ...sheet,
            [id]: tableStyle.style,
          });
        }
      }
    }
    this.host.apply(tr);
    return {
      docVersion: this.host.getVersion(),
      range: { from: insertAt, to: insertAt + inserted },
    };
  }

  async applyFormatting(
    target: FormatTarget,
    format: Formatting,
    opts: MutationOptions = {},
  ): Promise<MutationResult> {
    this.checkVersion(opts.expectedVersion);
    const hit = this.formatHit(target, opts.occurrence);
    const state = this.host.getState();
    const { schema } = state;
    let tr = state.tr;
    for (const [flag, markName] of Object.entries(MARK_BY_FLAG)) {
      const want = format[flag as keyof typeof MARK_BY_FLAG];
      if (want === undefined) continue;
      const mark = schema.marks[markName];
      if (!mark) continue;
      tr = want
        ? tr.addMark(hit.from, hit.to, mark.create())
        : tr.removeMark(hit.from, hit.to, mark);
    }
    if (
      format.fontSize !== undefined &&
      schema.marks['fontSize'] &&
      hit.to > hit.from
    ) {
      tr = tr.addMark(
        hit.from,
        hit.to,
        schema.marks['fontSize'].create({ size: format.fontSize }),
      );
    }
    const pAttrs: Record<string, unknown> = {};
    if (format.align) pAttrs['align'] = format.align;
    if (format.heading !== undefined) {
      pAttrs['heading'] = format.heading ? format.heading : null;
      if (format.heading) pAttrs['styleId'] = null;
    }
    if (format.style !== undefined) {
      pAttrs['styleId'] = format.style;
      if (format.style) pAttrs['heading'] = null;
    }
    if (format.tabs !== undefined) {
      const width = this.contentWidth();
      pAttrs['tabs'] = format.tabs.length
        ? format.tabs.map((t) => ({
            pos: Math.round(lengthToPx(t.at, width)),
            val: t.align ?? 'left',
            ...(t.leader ? { leader: t.leader } : {}),
          }))
        : null;
    }
    if (Object.keys(pAttrs).length > 0) {
      const block = this.textblocks()[hit.blockIndex];
      tr = tr.setNodeMarkup(block.pos, undefined, {
        ...block.node.attrs,
        ...pAttrs,
      });
    }
    if (tr.steps.length === 0) {
      // Formatting that named no supported change — a no-op success.
      return { docVersion: this.host.getVersion() };
    }
    this.host.apply(tr);
    return {
      docVersion: this.host.getVersion(),
      range: { from: hit.from, to: hit.to },
    };
  }

  async updateImage(
    blockIndex: number,
    imageIndex: number,
    changes: ImageChanges,
    opts: MutationOptions = {},
  ): Promise<MutationResult> {
    this.checkVersion(opts.expectedVersion);
    const blocks = this.textblocks();
    const block = blocks[blockIndex];
    if (!block) {
      throw new AnchorError(
        `blockIndex ${blockIndex} is out of range — the document has ${blocks.length} block(s). ` +
          `Block indexes change with every edit; call get_document again.`,
      );
    }
    const images = blockImages(block.node).map((img) => ({
      ...img,
      pos: block.pos + 1 + img.offset,
    }));
    if (images.length === 0) {
      throw new AnchorError(
        `Block ${blockIndex} has no images. get_document lists each block's images.`,
      );
    }
    const img = images[imageIndex];
    if (!img) {
      throw new AnchorError(
        `imageIndex ${imageIndex} is out of range — block ${blockIndex} has ${images.length} image(s) (0-${images.length - 1}).`,
      );
    }
    let tr = this.host.getState().tr;
    if (changes.width !== undefined)
      tr = tr.setNodeAttribute(
        img.pos,
        'width',
        Math.max(1, Math.round(changes.width)),
      );
    if (changes.height !== undefined)
      tr = tr.setNodeAttribute(
        img.pos,
        'height',
        Math.max(1, Math.round(changes.height)),
      );
    if (changes.rotation !== undefined) {
      tr = tr.setNodeAttribute(
        img.pos,
        'rotation',
        ((changes.rotation % 360) + 360) % 360,
      );
    }
    if (tr.steps.length === 0) return { docVersion: this.host.getVersion() };
    this.host.apply(tr);
    return {
      docVersion: this.host.getVersion(),
      range: { from: img.pos, to: img.pos + 1 },
    };
  }

  async save(): Promise<void> {
    await this.host.save();
  }

  async close(): Promise<void> {
    /* sessions over a host hold no resources of their own */
  }

  async editTable(
    tableIndex: number,
    edit: TableEdit,
    opts: MutationOptions = {},
  ): Promise<MutationResult & { rows: number; cols: number }> {
    this.checkVersion(opts.expectedVersion);
    const state = this.host.getState();
    const { schema } = state;
    const width = this.contentWidth();
    let tr = state.tr;
    const locate = () => {
      let found: { node: PMNode; pos: number } | null = null;
      let n = 0;
      tr.doc.descendants((node, pos) => {
        if (found) return false;
        if (node.type.name === 'table') {
          if (n === tableIndex) found = { node, pos };
          n++;
        }
        return !found;
      });
      if (!found) {
        throw new ContentError(
          `No table ${tableIndex} — get_document marks the blocks inside tables with table.index (0-based); the document has ${n}.`,
        );
      }
      return found as { node: PMNode; pos: number };
    };
    const rowAt = (table: PMNode, pos: number, r: number) => {
      if (r < 0 || r >= table.childCount) {
        throw new ContentError(
          `Row ${r} is out of range — the table has ${table.childCount} row(s).`,
        );
      }
      let rowPos = pos + 1;
      for (let i = 0; i < r; i++) rowPos += table.child(i).nodeSize;
      return { node: table.child(r), pos: rowPos };
    };

    if (edit.deleteRows?.length) {
      const { node, pos } = locate();
      const rows = [...new Set(edit.deleteRows)].sort((a, b) => b - a);
      if (rows.length >= node.childCount)
        throw new ContentError('A table must keep at least one row.');
      for (const r of rows) {
        const row = rowAt(node, pos, r);
        tr = tr.delete(row.pos, row.pos + row.node.nodeSize);
      }
    }
    if (edit.insertRows) {
      const { node, pos } = locate();
      const grid = tableGrid(node, width);
      const at = edit.insertRows.at ?? node.childCount;
      if (at < 0 || at > node.childCount) {
        throw new ContentError(
          `Cannot insert at row ${at} — the table has ${node.childCount} row(s); omit at to append.`,
        );
      }
      const rows = edit.insertRows.rows.map((cells) =>
        rowNode(cells, grid.widths, schema),
      );
      const insertAt =
        at === node.childCount
          ? pos + node.nodeSize - 1
          : rowAt(node, pos, at).pos;
      tr = tr.insert(insertAt, rows);
    }
    if (edit.merge) {
      const { node, pos } = locate();
      const { row: r, from, to } = edit.merge;
      const row = rowAt(node, pos, r);
      if (from < 0 || to >= row.node.childCount || from > to) {
        throw new ContentError(
          `merge cells ${from}..${to} is out of range — row ${r} has ${row.node.childCount} cell(s).`,
        );
      }
      if (from < to) {
        let cellPos = row.pos + 1;
        for (let i = 0; i < from; i++) cellPos += row.node.child(i).nodeSize;
        let end = cellPos;
        let colspan = 0;
        const colwidth: number[] = [];
        let content = row.node.child(from).content;
        for (let i = from; i <= to; i++) {
          const cell = row.node.child(i);
          end += cell.nodeSize;
          colspan += Math.max(1, Number(cell.attrs['colspan']) || 1);
          const cw = cell.attrs['colwidth'] as number[] | null;
          if (cw) colwidth.push(...cw);
          if (i > from) content = content.append(cell.content);
        }
        const first = row.node.child(from);
        const merged = first.type.create(
          {
            ...first.attrs,
            colspan,
            colwidth: colwidth.length === colspan ? colwidth : null,
          },
          content,
        );
        tr = tr.replaceWith(cellPos, end, merged);
      }
    }
    if (edit.widths) {
      const { node, pos } = locate();
      const grid = tableGrid(node, width);
      const widths = columnWidths(edit.widths, grid.cols, width);
      node.forEach((row, rowOffset) => {
        let col = 0;
        row.forEach((cell, cellOffset) => {
          const span = Math.max(1, Number(cell.attrs['colspan']) || 1);
          const cellPos = pos + 1 + rowOffset + 1 + cellOffset;
          tr = tr.setNodeMarkup(cellPos, undefined, {
            ...cell.attrs,
            colwidth: widths.slice(col, col + span),
          });
          col += span;
        });
      });
    }
    if (edit.borders || edit.align !== undefined || edit.header !== undefined) {
      const { node, pos } = locate();
      const attrs: Record<string, unknown> = { ...node.attrs };
      const tableStyle = this.host.tableStyle?.();
      if (edit.borders) {
        const chrome = tableChrome(
          edit.borders,
          tableStyle,
          !!node.type.spec.attrs?.['styleId'],
        );
        attrs['styleId'] = chrome.styleId;
        attrs['look'] = chrome.look;
        attrs['borders'] = chrome.borders;
        if (
          chrome.styleId &&
          tableStyle?.style &&
          schema.nodes['doc'].spec.attrs?.['tableStyles']
        ) {
          const sheet = (tr.doc.attrs['tableStyles'] ?? {}) as Record<
            string,
            unknown
          >;
          if (!sheet[chrome.styleId]) {
            tr = tr.setDocAttribute('tableStyles', {
              ...sheet,
              [chrome.styleId]: tableStyle.style,
            });
          }
        }
      }
      if (edit.align !== undefined)
        attrs['align'] = edit.align === 'left' ? null : edit.align;
      tr = tr.setNodeMarkup(pos, undefined, attrs);
      if (edit.header !== undefined) {
        const first = rowAt(locate().node, pos, 0);
        tr = tr.setNodeMarkup(first.pos, undefined, {
          ...first.node.attrs,
          header: edit.header,
        });
      }
    }
    if (tr.steps.length === 0) {
      const { node } = locate();
      return {
        docVersion: this.host.getVersion(),
        rows: node.childCount,
        cols: tableGrid(node, width).cols,
      };
    }
    this.host.apply(tr);
    const { node, pos } = locate();
    return {
      docVersion: this.host.getVersion(),
      range: { from: pos, to: pos + node.nodeSize },
      rows: node.childCount,
      cols: tableGrid(node, width).cols,
    };
  }

  // ── internals ────────────────────────────────────────────────────────

  /** Resolve a formatting target to absolute positions. */
  private formatHit(target: FormatTarget, occurrence?: number): Hit {
    if (typeof target === 'string') return this.uniqueHit(target, occurrence);
    const blocks = this.textblocks();
    const block = blocks[target.blockIndex];
    if (!block) {
      throw new AnchorError(
        `blockIndex ${target.blockIndex} is out of range — the document has ${blocks.length} block(s).`,
      );
    }
    return {
      from: block.pos + 1,
      to: block.pos + 1 + block.node.content.size,
      blockIndex: target.blockIndex,
      context: block.node.textContent,
    };
  }

  /** The text area's width in px, from the document's page setup. */
  private contentWidth(): number {
    const page = this.host.getState().doc.attrs['page'] as
      | {
          width: number;
          margin: { left: number; right: number };
          gutter?: number;
        }
      | null
      | undefined;
    if (!page) return DEFAULT_CONTENT_WIDTH;
    return (
      page.width - page.margin.left - page.margin.right - (page.gutter ?? 0)
    );
  }

  private checkVersion(expected?: string): void {
    const current = this.host.getVersion();
    if (expected !== undefined && expected !== current) {
      throw new VersionConflictError(current, expected);
    }
  }

  /** All textblocks (paragraphs, incl. inside table cells) in reading order,
   *  each knowing which table cell holds it. */
  private textblocks(): TextBlock[] {
    const out: TextBlock[] = [];
    let tables = 0;
    const walk = (node: PMNode, base: number, ctx: TextBlock['table']) => {
      node.forEach((child, offset, i) => {
        const pos = base + offset;
        if (child.isTextblock) {
          out.push(
            ctx ? { node: child, pos, table: ctx } : { node: child, pos },
          );
          return;
        }
        let next = ctx;
        if (child.type.name === 'table')
          next = { index: tables++, row: 0, cell: 0 };
        else if (child.type.name === 'table_row' && ctx)
          next = { ...ctx, row: i, cell: 0 };
        else if (child.type.name === 'table_cell' && ctx)
          next = { ...ctx, cell: i };
        walk(child, pos + 1, next);
      });
    };
    walk(this.host.getState().doc, 0, undefined);
    return out;
  }

  /** Every occurrence of `query`, atom-safe (matches never span images/fields
   *  or block boundaries), in document order with absolute PM positions. */
  private hits(query: string): Hit[] {
    if (query.length === 0) return [];
    const out: Hit[] = [];
    this.textblocks().forEach(({ node, pos }, blockIndex) => {
      // Concatenate the block's text children, breaking the searchable string
      // at non-text inlines so a match can't pretend to span an atom. Each
      // segment records where it starts in the joined string AND in the doc.
      const segments: {
        joinedStart: number;
        length: number;
        startPos: number;
      }[] = [];
      let joined = '';
      node.forEach((child, offset) => {
        if (child.isText && child.text) {
          segments.push({
            joinedStart: joined.length,
            length: child.text.length,
            startPos: pos + 1 + offset,
          });
          joined += child.text;
        } else {
          joined += '￿'; // unmatchable atom sentinel
        }
      });
      let at = joined.indexOf(query);
      while (at !== -1) {
        // A match may span adjacent text segments (mark changes split runs);
        // adjacency in `joined` implies adjacency in the doc, so mapping the
        // START offset to a position is enough.
        const seg = segments.find(
          (s) => at >= s.joinedStart && at < s.joinedStart + s.length,
        );
        if (seg) {
          const from = seg.startPos + (at - seg.joinedStart);
          out.push({
            from,
            to: from + query.length,
            blockIndex,
            context: contextAround(joined.replace(/￿/g, ' '), at, query.length),
          });
        }
        at = joined.indexOf(query, at + 1);
      }
    });
    return out;
  }

  private uniqueHit(text: string, occurrence?: number): Hit {
    const all = this.hits(text);
    if (all.length === 0) {
      throw new AnchorError(
        `Text not found in the document: ${JSON.stringify(clip(text))}. ` +
          `Anchors match within one paragraph — check get_document for the exact text.`,
      );
    }
    if (occurrence !== undefined) {
      const hit = all[occurrence - 1];
      if (!hit) {
        throw new AnchorError(
          `occurrence ${occurrence} is out of range — the text matches ${all.length} time(s).`,
        );
      }
      return hit;
    }
    if (all.length > 1) {
      throw new AnchorError(
        `The text matches ${all.length} times — pass occurrence (1-${all.length}) to pick one, ` +
          `or use a longer, unique anchor.`,
      );
    }
    return all[0];
  }
}

/** The block's inline image children, in order, with their child offsets. */
function blockImages(block: PMNode): { node: PMNode; offset: number }[] {
  const out: { node: PMNode; offset: number }[] = [];
  block.forEach((child, offset) => {
    if (child.type.name === 'image') out.push({ node: child, offset });
  });
  return out;
}

function blockType(node: PMNode): string {
  const heading = node.attrs['heading'] as number | null | undefined;
  if (typeof heading === 'number' && heading >= 1) return `heading${heading}`;
  return node.type.name;
}

function contextAround(text: string, at: number, len: number): string {
  const before = text.slice(Math.max(0, at - 30), at);
  const after = text.slice(at + len, at + len + 30);
  return `${before}«${text.slice(at, at + len)}»${after}`;
}

function clip(s: string): string {
  return s.length > 60 ? `${s.slice(0, 57)}…` : s;
}
