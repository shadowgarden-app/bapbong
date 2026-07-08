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
}

/** A text hit resolved to absolute PM positions. */
interface Hit {
  from: number;
  to: number;
  blockIndex: number;
  context: string;
}

const MARK_BY_FLAG = { bold: 'strong', italic: 'em', underline: 'underline', strike: 'strike' } as const;

export class PmDocSession implements DocumentSession {
  readonly capabilities: SessionCapabilities;

  constructor(private readonly host: PmSessionHost) {
    this.capabilities = { selection: typeof host.selection === 'function' };
  }

  // ── reads ────────────────────────────────────────────────────────────

  async snapshot(): Promise<DocSnapshot> {
    const blocks: DocBlock[] = this.textblocks().map(({ node }, index) => ({
      index,
      type: blockType(node),
      text: node.textContent,
    }));
    return { docVersion: this.host.getVersion(), blocks, meta: this.host.meta() };
  }

  async find(query: string): Promise<FindMatch[]> {
    return this.hits(query).map((h, i) => ({ blockIndex: h.blockIndex, occurrence: i + 1, context: h.context }));
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

  async replaceText(oldText: string, newText: string, opts: MutationOptions = {}): Promise<MutationResult> {
    this.checkVersion(opts.expectedVersion);
    const hit = this.uniqueHit(oldText, opts.occurrence);
    this.host.apply(this.host.getState().tr.insertText(newText, hit.from, hit.to));
    return { docVersion: this.host.getVersion() };
  }

  async insertContent(content: string, anchor: InsertAnchor, opts: MutationOptions = {}): Promise<MutationResult> {
    this.checkVersion(opts.expectedVersion);
    const state = this.host.getState();
    const { schema } = state;
    const paragraphs = content
      .split('\n')
      .map((line) => schema.node('paragraph', null, line.length > 0 ? [schema.text(line)] : []));

    let insertAt: number;
    if (anchor.position === 'document_end') {
      insertAt = state.doc.content.size;
    } else {
      const hit = this.uniqueHit(anchor.text, anchor.occurrence);
      const block = this.textblocks()[hit.blockIndex];
      insertAt = anchor.position === 'before' ? block.pos : block.pos + block.node.nodeSize;
    }
    this.host.apply(state.tr.insert(insertAt, paragraphs));
    return { docVersion: this.host.getVersion() };
  }

  async applyFormatting(target: string, format: Formatting, opts: MutationOptions = {}): Promise<MutationResult> {
    this.checkVersion(opts.expectedVersion);
    const hit = this.uniqueHit(target, opts.occurrence);
    const state = this.host.getState();
    const { schema } = state;
    let tr = state.tr;
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
      return { docVersion: this.host.getVersion() };
    }
    this.host.apply(tr);
    return { docVersion: this.host.getVersion() };
  }

  async save(): Promise<void> {
    await this.host.save();
  }

  async close(): Promise<void> {
    /* sessions over a host hold no resources of their own */
  }

  // ── internals ────────────────────────────────────────────────────────

  private checkVersion(expected?: string): void {
    const current = this.host.getVersion();
    if (expected !== undefined && expected !== current) {
      throw new VersionConflictError(current, expected);
    }
  }

  /** All textblocks (paragraphs, incl. inside table cells) in reading order. */
  private textblocks(): { node: PMNode; pos: number }[] {
    const out: { node: PMNode; pos: number }[] = [];
    this.host.getState().doc.descendants((node, pos) => {
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
        // A match may span adjacent text segments (mark changes split runs);
        // adjacency in `joined` implies adjacency in the doc, so mapping the
        // START offset to a position is enough.
        const seg = segments.find((s) => at >= s.joinedStart && at < s.joinedStart + s.length);
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
