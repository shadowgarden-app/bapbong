import { Component, ElementRef, signal, viewChild } from '@angular/core';
import { DOMSerializer, Node as ProseMirrorNode } from 'prosemirror-model';
import { schema } from '@shadow-garden/bapbong-model';
import { importDocx } from '@shadow-garden/bapbong-docx';
import { layout } from '@shadow-garden/bapbong-layout-engine';
import { createCanvasMeasurer, createCanvasMetrics } from '@shadow-garden/bapbong-measuring';
import { CanvasPainter } from '@shadow-garden/bapbong-painter-canvas';
import type { MeasureMetrics, MeasureText, PageConfig } from '@shadow-garden/bapbong-contracts';

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
      this.json.set(JSON.stringify(doc.toJSON(), null, 2));
      this.renderCanvas(doc);
      this.renderPreview(doc);
    } catch (err) {
      this.error.set(err instanceof Error ? err.message : String(err));
    } finally {
      this.loading.set(false);
    }
  }

  /** M3: lay the document out and paint it onto the canvas. */
  private renderCanvas(doc: ProseMirrorNode): void {
    const canvas = this.canvasHost()?.nativeElement;
    if (!canvas) return;
    this.measureText ??= createCanvasMeasurer();
    this.measureMetrics ??= createCanvasMetrics();
    const resolved = layout(doc, {
      page: A4,
      measureText: this.measureText,
      measureMetrics: this.measureMetrics,
    });
    this.painter ??= new CanvasPainter(canvas);
    this.painter.paint(resolved);
    this.pageCount.set(resolved.pages.length);
  }

  /** Render the document with the schema's own toDOM rules. */
  private renderPreview(doc: ProseMirrorNode): void {
    const host = this.previewHost()?.nativeElement;
    if (!host) return;
    host.replaceChildren(this.serializer.serializeFragment(doc.content, { document }));
  }
}
