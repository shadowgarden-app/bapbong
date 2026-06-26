import { DOMSerializer, type Node as ProseMirrorNode } from 'prosemirror-model';

export interface A11yMirrorOptions {
  /** `aria-label` for the mirrored document region. */
  label?: string;
  /** Debounce (ms) before re-serializing after `update()` — coalesces rapid
   *  edits so typing stays snappy (the mirror is for reading/navigation, not
   *  live keystroke echo). Default 200. Set 0 to update synchronously. */
  debounceMs?: number;
}

/** Visually-hidden but assistive-tech-readable. The standard "sr-only" recipe:
 *  clipped to a 1px box, off the visual layer, yet present in the a11y tree. */
const SR_ONLY =
  'position:absolute;width:1px;height:1px;margin:-1px;padding:0;border:0;overflow:hidden;clip:rect(0 0 0 0);clip-path:inset(50%);white-space:normal;';

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
  private pending: ProseMirrorNode | null = null;
  private timer: ReturnType<typeof setTimeout> | null = null;

  constructor(host: HTMLElement, opts: A11yMirrorOptions = {}) {
    this.debounceMs = opts.debounceMs ?? 200;
    const el = host.ownerDocument.createElement('div');
    el.setAttribute('role', 'document');
    el.setAttribute('aria-label', opts.label ?? 'Document content');
    el.style.cssText = SR_ONLY;
    host.appendChild(el);
    this.el = el;
  }

  /** Re-render the mirror from `doc` (debounced). Safe to call on every change. */
  update(doc: ProseMirrorNode): void {
    this.pending = doc;
    if (this.debounceMs <= 0) {
      this.flush();
      return;
    }
    if (this.timer != null) return; // a flush is already scheduled
    this.timer = setTimeout(() => {
      this.timer = null;
      this.flush();
    }, this.debounceMs);
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
    const serializer = DOMSerializer.fromSchema(doc.type.schema);
    const fragment = serializer.serializeFragment(doc.content, {
      document: this.el.ownerDocument,
    });
    this.el.replaceChildren(fragment);
    // Drop the heavy base64 `src` on inline images — assistive tech reads the
    // `alt`, and duplicating data-URLs would bloat the mirror DOM.
    for (const img of Array.from(this.el.querySelectorAll('img'))) {
      img.removeAttribute('src');
      if (!img.getAttribute('alt')) img.setAttribute('alt', 'image');
    }
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
