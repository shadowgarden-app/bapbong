import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { RenderCore } from './render-core.js';

/**
 * Read-only text selection for {@link BapbongView}. A canvas paints text as
 * pixels — there is nothing to drag-select — so this mirrors the editor's
 * approach: hit-test the pointer to document positions, paint the selection
 * rects on the canvas overlay (same geometry the painter uses, so it aligns
 * exactly at any zoom), and copy the selected text from the document model.
 *
 * No input-bridge / ProseMirror editing is involved — it only reads the doc and
 * the layout the {@link RenderCore} already holds. (The sr-only ARIA mirror is a
 * separate concern: it serves screen readers, not mouse selection.)
 */
export class ViewerSelection {
  private anchor: number | null = null;
  private head: number | null = null;
  private dragging = false;

  constructor(private readonly core: RenderCore) {
    const s = core.stack;
    s.style.cursor = 'text';
    // Focusable so Ctrl/Cmd+C (copy) and Ctrl/Cmd+A (select all) reach us.
    if (s.tabIndex < 0) s.tabIndex = 0;
    s.style.outline = 'none';
    s.addEventListener('pointerdown', this.onDown);
    s.addEventListener('pointermove', this.onMove);
    s.addEventListener('pointerup', this.onUp);
    s.addEventListener('pointercancel', this.onUp);
    s.addEventListener('dblclick', this.onDblClick);
    s.addEventListener('keydown', this.onKey);
  }

  /** Drop any current selection (e.g. when a new document loads). */
  clear(): void {
    this.anchor = this.head = null;
    this.dragging = false;
    this.core.paintOverlay({ selection: [] });
  }

  destroy(): void {
    const s = this.core.stack;
    s.removeEventListener('pointerdown', this.onDown);
    s.removeEventListener('pointermove', this.onMove);
    s.removeEventListener('pointerup', this.onUp);
    s.removeEventListener('pointercancel', this.onUp);
    s.removeEventListener('dblclick', this.onDblClick);
    s.removeEventListener('keydown', this.onKey);
  }

  // ── Pointer ─────────────────────────────────────────────────────────

  private onDown = (ev: PointerEvent): void => {
    if (ev.button !== 0) return;
    const pos = this.core.posAtEvent(ev);
    if (pos == null) return;
    ev.preventDefault(); // suppress native selection (e.g. of the sr-only mirror)
    this.core.stack.focus?.();
    this.anchor = this.head = pos;
    this.dragging = true;
    this.paint();
    try {
      (ev.target as HTMLElement).setPointerCapture?.(ev.pointerId);
    } catch {
      // pointer capture is a nicety
    }
  };

  private onMove = (ev: PointerEvent): void => {
    if (!this.dragging || this.anchor == null || !(ev.buttons & 1)) return;
    const coalesced = ev.getCoalescedEvents?.() ?? [];
    const last = coalesced.length > 0 ? coalesced[coalesced.length - 1] : ev;
    const pos = this.core.posAtEvent(last);
    if (pos == null || pos === this.head) return;
    this.head = pos;
    this.paint();
  };

  private onUp = (): void => {
    this.dragging = false;
  };

  /** Double-click selects the word under the pointer. */
  private onDblClick = (ev: MouseEvent): void => {
    const pos = this.core.posAtEvent(ev);
    const doc = this.core.document;
    if (pos == null || !doc) return;
    const word = wordRangeAt(doc, pos);
    if (!word) return;
    ev.preventDefault();
    this.anchor = word.from;
    this.head = word.to;
    this.paint();
  };

  // ── Keyboard ────────────────────────────────────────────────────────

  private onKey = (ev: KeyboardEvent): void => {
    if (!(ev.metaKey || ev.ctrlKey)) return;
    const key = ev.key.toLowerCase();
    if (key === 'c') {
      const text = this.selectedText();
      if (text) {
        ev.preventDefault();
        navigator.clipboard?.writeText(text).catch(() => {
          /* clipboard unavailable */
        });
      }
    } else if (key === 'a') {
      const doc = this.core.document;
      if (!doc) return;
      ev.preventDefault();
      this.anchor = 0;
      this.head = doc.content.size;
      this.paint();
    }
  };

  // ── Helpers ─────────────────────────────────────────────────────────

  private range(): { from: number; to: number } | null {
    if (this.anchor == null || this.head == null || this.anchor === this.head) return null;
    return { from: Math.min(this.anchor, this.head), to: Math.max(this.anchor, this.head) };
  }

  private paint(): void {
    const r = this.range();
    this.core.paintOverlay({ selection: r ? this.core.selectionRects(r.from, r.to) : [] });
  }

  private selectedText(): string | null {
    const r = this.range();
    const doc = this.core.document;
    if (!r || !doc) return null;
    return doc.textBetween(r.from, r.to, '\n', '\n');
  }
}

/** The word range (doc positions) around `pos`, or null. Scans the text of the
 *  containing block; treats runs of non-space as words. */
function wordRangeAt(doc: ProseMirrorNode, pos: number): { from: number; to: number } | null {
  const $pos = doc.resolve(pos);
  const text = $pos.parent.textContent;
  if (!text) return null;
  const start = $pos.start();
  let off = pos - start;
  off = Math.max(0, Math.min(off, text.length));
  const isWord = (c: string) => c != null && !/\s/.test(c);
  // Anchor onto a word char (prefer the char before the caret, then after).
  if (!isWord(text[off]) && isWord(text[off - 1])) off -= 1;
  if (!isWord(text[off])) return null;
  let from = off;
  let to = off;
  while (from > 0 && isWord(text[from - 1])) from--;
  while (to < text.length && isWord(text[to])) to++;
  return { from: start + from, to: start + to };
}
