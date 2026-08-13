import type { Schema } from 'prosemirror-model';
import { RenderCore, type RenderCoreOptions } from './render-core.js';
import { ViewerSelection } from './viewer-selection.js';

export interface BapbongViewOptions extends RenderCoreOptions {
  /** The scroll viewport for virtualization / scroll-into-view. Defaults to the
   *  host element itself (give it `overflow:auto` + a bounded height). */
  viewport?: HTMLElement;
  /** Allow read-only text selection + copy by dragging over the canvas
   *  (double-click selects a word, Ctrl/⌘+A selects all, Ctrl/⌘+C copies).
   *  Default true; pass false for a non-interactive preview. */
  selectable?: boolean;
}

/** Emitted after the document (re)renders — page count changed, fonts settled. */
export interface ViewChange {
  pageCount: number;
}

/**
 * A **read-only** bapbong document viewer (preview). Give it a scrollable host
 * element; it renders the imported `.docx` to a virtualized canvas page-stack
 * with zoom — no editing, no input-bridge, so the preview bundle never pulls in
 * the ProseMirror editing surface.
 *
 *   const view = new BapbongView(hostEl);   // host: overflow:auto, fixed height
 *   await view.loadDocx(bytes);
 *   view.setZoom(1.25);
 *
 * For an editable instance use `@shadow-garden/bapbong-editor` (it composes the
 * same {@link RenderCore}). For headless/server use `@shadow-garden/bapbong-headless`.
 */
export class BapbongView {
  private readonly stack: HTMLElement;
  private readonly core: RenderCore;
  private readonly selection: ViewerSelection | null;
  private readonly changeListeners = new Set<(c: ViewChange) => void>();
  private readonly offFonts: () => void;

  constructor(host: HTMLElement, opts: BapbongViewOptions = {}) {
    // The page canvases stack inside a relatively-positioned child so the host
    // stays the scroll container (virtualization measures against it).
    const stack = document.createElement('div');
    stack.style.position = 'relative';
    host.appendChild(stack);
    this.stack = stack;
    // Spread the caller's options rather than copying four of them: this used
    // to name each field, which silently dropped measureText/measureMetrics —
    // an embedder that supplied real font metrics still got canvas ones, and
    // the same document paginated differently in the viewer and the editor.
    this.core = new RenderCore(stack, {
      ...opts,
      viewport: opts.viewport ?? host,
    });
    this.offFonts = this.core.onFontsReloaded(() => this.emit());
    this.selection =
      opts.selectable === false ? null : new ViewerSelection(this.core);
  }

  /** Import a `.docx` and render the first frame. Resolves with the imported
   *  page-chrome keys (header/footer w:types present). */
  async loadDocx(
    bytes: ArrayBuffer,
  ): Promise<{ headerKeys: string[]; footerKeys: string[] }> {
    this.selection?.clear();
    const { headerKeys, footerKeys } = await this.core.loadDocx(bytes);
    this.emit();
    return { headerKeys, footerKeys };
  }

  /** Export the current document back to .docx bytes (carries the source
   *  package so unmodelled parts survive the round-trip). */
  exportDocx(): Promise<Uint8Array> {
    return this.core.exportDocx();
  }

  /** Scroll so the caret at doc position `pos` sits `topMargin` px from the top. */
  scrollToPos(pos: number, topMargin?: number): void {
    this.core.scrollToPos(pos, topMargin);
  }

  /** Set the zoom factor (1 = 100%) and repaint at the new scale. */
  setZoom(zoom: number): void {
    this.core.setZoom(zoom);
  }

  /** The current zoom factor. */
  getZoom(): number {
    return this.core.getZoom();
  }

  /** Print the whole document (renders every page, one per sheet). */
  print(): Promise<void> {
    return this.core.print();
  }

  /** Number of laid-out pages (0 before the first document). */
  get pageCount(): number {
    return this.core.pageCount;
  }

  /** The document schema in use. */
  get schema(): Schema {
    return this.core.schema;
  }

  /** Subscribe to render cycles (document loaded, page count changed). Returns
   *  an unsubscribe fn. */
  onChange(cb: (c: ViewChange) => void): () => void {
    this.changeListeners.add(cb);
    return () => this.changeListeners.delete(cb);
  }

  destroy(): void {
    this.offFonts();
    this.selection?.destroy();
    this.core.destroy();
    this.stack.remove();
  }

  private emit(): void {
    const change: ViewChange = { pageCount: this.core.pageCount };
    for (const cb of this.changeListeners) cb(change);
  }
}
