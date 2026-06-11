import { Component, ElementRef, signal, viewChild } from '@angular/core';
import { DOMSerializer, Node as ProseMirrorNode } from 'prosemirror-model';
import { schema } from '@shadow-garden/bapbong-model';
import { importDocx } from '@shadow-garden/bapbong-docx';
import { layout } from '@shadow-garden/bapbong-layout-engine';
import { createCanvasMeasurer, createCanvasMetrics } from '@shadow-garden/bapbong-measuring';
import { CanvasPainter } from '@shadow-garden/bapbong-painter-canvas';
import {
  InputBridge,
  moveCaretCommand,
  type Command,
  type EditorState,
} from '@shadow-garden/bapbong-input-bridge';
import { caretRect, hitTest, selectionRects, verticalCaret } from '@shadow-garden/bapbong-selection';
import type {
  MeasureMetrics,
  MeasureText,
  PageConfig,
  ResolvedLayout,
} from '@shadow-garden/bapbong-contracts';

/** A4 at 96 dpi with 1in margins. */
const A4: PageConfig = {
  width: 794,
  height: 1123,
  margin: { top: 96, right: 96, bottom: 96, left: 96 },
};

@Component({
  selector: 'app-root',
  templateUrl: './app.html',
  styleUrl: './app.css',
})
export class App {
  protected readonly fileName = signal<string | null>(null);
  protected readonly error = signal<string | null>(null);
  protected readonly json = signal<string | null>(null);
  protected readonly headerKeys = signal<string[]>([]);
  protected readonly footerKeys = signal<string[]>([]);
  protected readonly loading = signal(false);
  protected readonly pageCount = signal(0);

  private readonly previewHost = viewChild<ElementRef<HTMLDivElement>>('preview');
  private readonly canvasHost = viewChild<ElementRef<HTMLCanvasElement>>('docCanvas');
  private readonly serializer = DOMSerializer.fromSchema(schema);

  private painter: CanvasPainter | null = null;
  private measureText: MeasureText | null = null;
  private measureMetrics: MeasureMetrics | null = null;
  private bridge: InputBridge | null = null;
  private resolved: ResolvedLayout | null = null;
  private dragAnchor: number | null = null;

  protected async onFile(event: Event): Promise<void> {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (!file) return;
    await this.load(file.name, await file.arrayBuffer());
  }

  protected async loadSample(): Promise<void> {
    try {
      const res = await fetch('sample.docx');
      if (!res.ok) throw new Error(`sample.docx: HTTP ${res.status}`);
      await this.load('sample.docx', await res.arrayBuffer());
    } catch (err) {
      this.error.set(err instanceof Error ? err.message : String(err));
    }
  }

  private async load(name: string, bytes: ArrayBuffer): Promise<void> {
    this.loading.set(true);
    this.error.set(null);
    this.json.set(null);
    this.fileName.set(name);

    try {
      const { doc, headers, footers } = await importDocx(bytes);
      this.headerKeys.set(Object.keys(headers));
      this.footerKeys.set(Object.keys(footers));
      this.setupEditor(doc);
    } catch (err) {
      this.error.set(err instanceof Error ? err.message : String(err));
    } finally {
      this.loading.set(false);
    }
  }

  /** M4: mount the hidden ProseMirror editor and run the first paint. */
  private setupEditor(doc: ProseMirrorNode): void {
    const canvas = this.canvasHost()?.nativeElement;
    if (!canvas) return;
    this.measureText ??= createCanvasMeasurer();
    this.measureMetrics ??= createCanvasMetrics();
    this.painter ??= new CanvasPainter(canvas);

    this.bridge?.destroy();
    this.bridge = new InputBridge({
      doc,
      keys: { ArrowUp: this.verticalCmd(-1), ArrowDown: this.verticalCmd(1) },
      onUpdate: (state) => this.refresh(state),
    });
    // Same scroll container as the canvas, so IME anchoring scrolls along.
    canvas.parentElement?.appendChild(this.bridge.dom);
    this.refresh(this.bridge.state);
  }

  /** Layout → paint (with caret/selection) → sync side panels. */
  private refresh(state: EditorState): void {
    if (!this.painter || !this.measureText || !this.measureMetrics) return;
    this.resolved = layout(state.doc, {
      page: A4,
      measureText: this.measureText,
      measureMetrics: this.measureMetrics,
    });
    const sel = state.selection;
    const caret = caretRect(this.resolved, sel.head, this.measureText);
    const rects = sel.empty ? [] : selectionRects(this.resolved, sel.from, sel.to, this.measureText);
    this.painter.paint(this.resolved, { caret, selection: rects });
    this.pageCount.set(this.resolved.pages.length);
    this.json.set(JSON.stringify(state.doc.toJSON(), null, 2));
    this.renderPreview(state.doc);

    // Anchor the hidden editor (and its IME popup) at the painted caret.
    if (caret && this.bridge) {
      const canvas = this.canvasHost()?.nativeElement;
      const pt = this.painter.pageToCanvas({ pageIndex: caret.pageIndex, x: caret.x, y: caret.y });
      if (canvas && pt) {
        this.bridge.place(canvas.offsetLeft + pt.x, canvas.offsetTop + pt.y, caret.height);
      }
    }
  }

  /** ArrowUp/ArrowDown against the canvas layout (the hidden DOM's own line
   *  wrapping is meaningless). */
  private verticalCmd(dir: -1 | 1): Command {
    return moveCaretCommand((state) => {
      if (!this.resolved || !this.measureText) return null;
      const head = state.selection.head;
      const cr = caretRect(this.resolved, head, this.measureText);
      if (!cr) return null;
      return verticalCaret(this.resolved, head, dir, cr.x, this.measureText);
    });
  }

  protected onCanvasMouseDown(ev: MouseEvent): void {
    const pos = this.posAtEvent(ev);
    if (pos == null || !this.bridge) return;
    ev.preventDefault(); // keep focus on the hidden editor
    this.dragAnchor = pos;
    this.bridge.setSelection(pos);
    this.bridge.focus();
  }

  protected onCanvasMouseMove(ev: MouseEvent): void {
    if (this.dragAnchor == null || !(ev.buttons & 1)) return;
    const pos = this.posAtEvent(ev);
    if (pos == null || !this.bridge) return;
    this.bridge.setSelection(this.dragAnchor, pos);
  }

  protected onCanvasMouseUp(): void {
    this.dragAnchor = null;
  }

  protected onCanvasDblClick(ev: MouseEvent): void {
    const pos = this.posAtEvent(ev);
    if (pos == null || !this.bridge) return;
    ev.preventDefault();
    this.bridge.selectWordAt(pos);
    this.bridge.focus();
  }

  private posAtEvent(ev: MouseEvent): number | null {
    if (!this.painter || !this.resolved || !this.measureText) return null;
    const pt = this.painter.canvasToPage(ev.offsetX, ev.offsetY);
    if (!pt) return null;
    return hitTest(this.resolved, pt, this.measureText);
  }

  /** Render the document with the schema's own toDOM rules. */
  private renderPreview(doc: ProseMirrorNode): void {
    const host = this.previewHost()?.nativeElement;
    if (!host) return;
    host.replaceChildren(this.serializer.serializeFragment(doc.content, { document }));
  }
}
