/**
 * DocumentSession over a headless ProseMirror EditorState — no DOM, no UI.
 *
 * This is both the unit-test host for the MCP tool semantics and the seed of
 * a future server-side document service (a FastAPI-style gateway keeps docs
 * open in one of these per session). The desktop app implements the same
 * port against its live editor instead.
 */
import { importDocx, exportDocx, type DocxImport } from '@shadow-garden/bapbong-headless';
import { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import {
  AnchorError,
  VersionConflictError,
  type DocBlock,
  type DocSnapshot,
  type DocumentSession,
  type FindMatch,
  type Formatting,
  type InsertAnchor,
  type MutationOptions,
  type MutationResult,
  type SessionCapabilities,
} from './contract.js';

/** A text hit resolved to absolute PM positions. */
interface Hit {
  from: number;
  to: number;
  blockIndex: number;
  context: string;
}

const MARK_BY_FLAG = { bold: 'strong', italic: 'em', underline: 'underline', strike: 'strike' } as const;

export interface HeadlessSessionOptions {
  /** Where save() writes the exported bytes (file, DB, test sink…). */
  onSave?: (bytes: Uint8Array) => void | Promise<void>;
  name?: string;
}

export class HeadlessSession implements DocumentSession {
  readonly capabilities: SessionCapabilities = { selection: false };

  private state: EditorState;
  private version = 1;
  private dirty = false;
  private readonly raw: DocxImport['raw'];
  private readonly opts: HeadlessSessionOptions;

  private constructor(state: EditorState, raw: DocxImport['raw'], opts: HeadlessSessionOptions) {
    this.state = state;
    this.raw = raw;
    this.opts = opts;
  }

  static async open(bytes: ArrayBuffer | Uint8Array, opts: HeadlessSessionOptions = {}): Promise<HeadlessSession> {
    const { doc, raw } = await importDocx(bytes instanceof Uint8Array ? (bytes.slice().buffer as ArrayBuffer) : bytes);
    return new HeadlessSession(EditorState.create({ doc }), raw, opts);
  }

  get docVersion(): string {
    return `v${this.version}`;
  }

  // ── reads ────────────────────────────────────────────────────────────

  async snapshot(): Promise<DocSnapshot> {
    const blocks: DocBlock[] = this.textblocks().map(({ node }, index) => ({
      index,
      type: blockType(node),
      text: node.textContent,
    }));
    return {
      docVersion: this.docVersion,
      blocks,
      meta: { name: this.opts.name, dirty: this.dirty },
    };
  }

  async find(query: string): Promise<FindMatch[]> {
    return this.hits(query).map((h, i) => ({ blockIndex: h.blockIndex, occurrence: i + 1, context: h.context }));
  }

  // ── mutations ────────────────────────────────────────────────────────

  async replaceText(oldText: string, newText: string, opts: MutationOptions = {}): Promise<MutationResult> {
    this.checkVersion(opts.expectedVersion);
    const hit = this.uniqueHit(oldText, opts.occurrence);
    const tr = this.state.tr.insertText(newText, hit.from, hit.to);
    return this.apply(tr);
  }

  async insertContent(content: string, anchor: InsertAnchor, opts: MutationOptions = {}): Promise<MutationResult> {
    this.checkVersion(opts.expectedVersion);
    const { schema } = this.state;
    const paragraphs = content
      .split('\n')
      .map((line) => schema.node('paragraph', null, line.length > 0 ? [schema.text(line)] : []));

    let insertAt: number;
    if (anchor.position === 'document_end') {
      insertAt = this.state.doc.content.size;
    } else {
      const hit = this.uniqueHit(anchor.text, anchor.occurrence);
      const block = this.textblocks()[hit.blockIndex];
      insertAt = anchor.position === 'before' ? block.pos : block.pos + block.node.nodeSize;
    }
    const tr = this.state.tr.insert(insertAt, paragraphs);
    return this.apply(tr);
  }

  async applyFormatting(target: string, format: Formatting, opts: MutationOptions = {}): Promise<MutationResult> {
    this.checkVersion(opts.expectedVersion);
    const hit = this.uniqueHit(target, opts.occurrence);
    const { schema } = this.state;
    let tr = this.state.tr;
    for (const [flag, markName] of Object.entries(MARK_BY_FLAG)) {
      const want = format[flag as keyof typeof MARK_BY_FLAG];
      if (want === undefined) continue;
      const mark = schema.marks[markName];
      if (!mark) continue;
      tr = want ? tr.addMark(hit.from, hit.to, mark.create()) : tr.removeMark(hit.from, hit.to, mark);
    }
    if (format.align) {
      const block = this.textblocks()[hit.blockIndex];
      tr = tr.setNodeMarkup(block.pos, undefined, { ...block.node.attrs, align: format.align });
    }
    if (tr.steps.length === 0) {
      // Formatting that named no supported change — a no-op success.
      return { docVersion: this.docVersion };
    }
    return this.apply(tr);
  }

  async save(): Promise<void> {
    const bytes = await exportDocx(this.state.doc, this.raw ? { carry: this.raw } : undefined);
    await this.opts.onSave?.(bytes);
    this.dirty = false;
  }

  async close(): Promise<void> {
    /* headless sessions hold no external resources */
  }

  // ── internals ────────────────────────────────────────────────────────

  private apply(tr: Parameters<EditorState['apply']>[0]): MutationResult {
    this.state = this.state.apply(tr);
    this.version++;
    this.dirty = true;
    return { docVersion: this.docVersion };
  }

  private checkVersion(expected?: string): void {
    if (expected !== undefined && expected !== this.docVersion) {
      throw new VersionConflictError(this.docVersion, expected);
    }
  }

  /** All textblocks (paragraphs, incl. inside table cells) in reading order. */
  private textblocks(): { node: PMNode; pos: number }[] {
    const out: { node: PMNode; pos: number }[] = [];
    this.state.doc.descendants((node, pos) => {
      if (node.isTextblock) {
        out.push({ node, pos });
        return false;
      }
      return true;
    });
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
      const segments: { joinedStart: number; length: number; startPos: number }[] = [];
      let joined = '';
      node.forEach((child, offset) => {
        if (child.isText && child.text) {
          segments.push({ joinedStart: joined.length, length: child.text.length, startPos: pos + 1 + offset });
          joined += child.text;
        } else {
          joined += '￿'; // unmatchable atom sentinel
        }
      });
      let at = joined.indexOf(query);
      while (at !== -1) {
        const seg = segments.find((s) => at >= s.joinedStart && at + query.length <= s.joinedStart + s.length)
          ?? segments.find((s) => at >= s.joinedStart && at < s.joinedStart + s.length);
        // A match crossing segment boundaries is fine ONLY between adjacent
        // text segments (mark changes split runs); adjacency in `joined`
        // implies adjacency in the doc, so from-position math still holds.
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
        throw new AnchorError(`occurrence ${occurrence} is out of range — the text matches ${all.length} time(s).`);
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
