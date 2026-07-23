import { DOMSerializer, type Node as ProseMirrorNode } from 'prosemirror-model';
import { perf } from '@shadow-garden/bapbong-contracts';

export interface A11yMirrorOptions {
  /** `aria-label` for the mirrored document region. */
  label?: string;
  /** Idle delay (ms): re-serialize this long after the last `update()`. The
   *  timer resets on each edit, so a rapid burst produces no serialize until it
   *  settles — the whole-document serialize (tens of ms on a large doc) never
   *  competes with a keystroke. Default 300. Set 0 to update synchronously. */
  debounceMs?: number;
  /** Hard cap (ms) on staleness: even during sustained typing (where the idle
   *  debounce would never fire) the mirror re-serializes at least this often.
   *  Default 2000. */
  maxWaitMs?: number;
  /** Documents with more top-level blocks than this render as a heading
   *  OUTLINE instead of a full mirror. A whole-document hidden DOM stalls
   *  WebKit's rendering pipeline on large docs (frames stop for seconds after
   *  every scroll — ablation-verified), so beyond this size screen readers get
   *  the navigable structure rather than the full text. Default 300. */
  maxFullBlocks?: number;
}

/** Visually-hidden but assistive-tech-readable. The standard "sr-only" recipe:
 *  clipped to a 1px box, off the visual layer, yet present in the a11y tree. */
const SR_ONLY =
  'position:absolute;width:1px;height:1px;margin:-1px;padding:0;border:0;overflow:hidden;clip:rect(0 0 0 0);clip-path:inset(50%);white-space:normal;contain:strict;';


/**
 * An **ARIA shadow-DOM mirror** of a bapbong document. The canvas the editor /
 * viewer paints into is opaque to screen readers; this keeps a parallel,
 * visually-hidden DOM subtree — built from the schema's own `toDOM` via
 * {@link DOMSerializer} — so assistive tech can read and navigate the content
 * (headings/paragraphs/lists/tables and real `<a href>` links).
 *
 *   const mirror = new A11yMirror(hostEl);
 *   mirror.update(doc);   // re-serialise on each doc change (debounced)
 *   mirror.destroy();
 *
 * It only needs the document `Node` (the schema rides on `doc.type.schema`), so
 * it works for both the editable and read-only tiers.
 */
export class A11yMirror {
  private readonly el: HTMLElement;
  private readonly debounceMs: number;
  private readonly maxWaitMs: number;
  private readonly maxFullBlocks: number;
  private pending: ProseMirrorNode | null = null;
  private timer: ReturnType<typeof setTimeout> | null = null;
  private lastFlushAt: number;

  constructor(host: HTMLElement, opts: A11yMirrorOptions = {}) {
    this.maxFullBlocks = opts.maxFullBlocks ?? 300;
    // The mirror is read by assistive tech when the user pauses to navigate, not
    // while typing (that goes through the hidden editor), so it can lag well
    // behind live edits. On a large document each re-serialize rebuilds the
    // whole hidden DOM — a big allocate-and-discard — so the defaults are
    // deliberately lazy: skip typing entirely, refresh a beat after a pause, and
    // cap staleness generously during sustained typing.
    this.debounceMs = opts.debounceMs ?? 1000;
    this.maxWaitMs = opts.maxWaitMs ?? 5000;
    this.lastFlushAt = perf.now();
    const el = host.ownerDocument.createElement('div');
    el.className = 'bapbong-a11y-mirror';
    el.setAttribute('role', 'document');
    el.setAttribute('aria-label', opts.label ?? 'Document content');
    el.style.cssText = SR_ONLY;
    host.appendChild(el);
    this.el = el;
  }

  /** Re-render the mirror from `doc`. Safe to call on every change: the
   *  whole-document serialize is debounced (reset on each edit) so a rapid
   *  keystroke burst produces none until it settles, and capped by `maxWaitMs`
   *  so the mirror still refreshes during sustained typing. */
  update(doc: ProseMirrorNode): void {
    this.pending = doc;
    if (this.debounceMs <= 0) {
      this.flush();
      return;
    }
    if (this.timer != null) clearTimeout(this.timer);
    // Idle debounce, but never let staleness exceed maxWaitMs since the last
    // serialize — as that cap approaches the delay shrinks toward 0.
    const sinceFlush = perf.now() - this.lastFlushAt;
    const delay = Math.min(this.debounceMs, Math.max(0, this.maxWaitMs - sinceFlush));
    this.timer = setTimeout(() => {
      this.timer = null;
      this.flush();
    }, delay);
  }

  /** Render any pending document immediately (cancelling a scheduled flush). */
  flush(): void {
    if (this.timer != null) {
      clearTimeout(this.timer);
      this.timer = null;
    }
    const doc = this.pending;
    if (!doc) return;
    this.pending = null;
    this.lastFlushAt = perf.now();
    if (doc.childCount > this.maxFullBlocks) {
      perf.span('a11y.flush(outline)', () => this.renderOutline(doc));
      return;
    }
    perf.span('a11y.flush', () => {
      const serializer = DOMSerializer.fromSchema(doc.type.schema);
      const fragment = serializer.serializeFragment(doc.content, {
        document: this.el.ownerDocument,
      });
      this.el.replaceChildren(fragment);
    });
    // Drop the heavy base64 `src` on inline images — assistive tech reads the
    // `alt`, and duplicating data-URLs would bloat the mirror DOM.
    for (const img of Array.from(this.el.querySelectorAll('img'))) {
      img.removeAttribute('src');
      if (!img.getAttribute('alt')) img.setAttribute('alt', 'image');
    }
  }

  /** Large-document mode: a navigable heading outline instead of the full
   *  mirror (see {@link A11yMirrorOptions.maxFullBlocks}). */
  private renderOutline(doc: ProseMirrorNode): void {
    const out = this.el.ownerDocument;
    const frag = out.createDocumentFragment();
    const note = out.createElement('p');
    note.textContent =
      'Large document: heading outline only. Full text is available in the editor.';
    frag.appendChild(note);
    let count = 0;
    doc.descendants((node) => {
      if (count >= 500) return false;
      if (!node.isTextblock) return true;
      const level = node.attrs['heading'] as number | null | undefined;
      if (typeof level === 'number' && level >= 1 && node.textContent.trim()) {
        const h = out.createElement(`h${Math.min(6, Math.max(1, level))}`);
        h.textContent = node.textContent.slice(0, 200);
        frag.appendChild(h);
        count++;
      }
      return false; // headings are top-level textblocks; no need to descend
    });
    this.el.replaceChildren(frag);
  }

  /** The mirror container (e.g. to wire `aria-describedby`). */
  get element(): HTMLElement {
    return this.el;
  }

  destroy(): void {
    if (this.timer != null) clearTimeout(this.timer);
    this.timer = null;
    this.pending = null;
    this.el.remove();
  }
}
