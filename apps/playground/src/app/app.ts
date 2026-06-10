import { Component, ElementRef, signal, viewChild } from '@angular/core';
import { DOMSerializer, Node as ProseMirrorNode } from 'prosemirror-model';
import { schema } from '@shadow-garden/bapbong-model';
import { importDocx } from '@shadow-garden/bapbong-docx';

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

  private readonly previewHost = viewChild<ElementRef<HTMLDivElement>>('preview');
  private readonly serializer = DOMSerializer.fromSchema(schema);

  protected async onFile(event: Event): Promise<void> {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (!file) return;

    this.loading.set(true);
    this.error.set(null);
    this.json.set(null);
    this.fileName.set(file.name);

    try {
      const { doc, headers, footers } = await importDocx(await file.arrayBuffer());
      this.headerKeys.set(Object.keys(headers));
      this.footerKeys.set(Object.keys(footers));
      this.json.set(JSON.stringify(doc.toJSON(), null, 2));
      this.renderPreview(doc);
    } catch (err) {
      this.error.set(err instanceof Error ? err.message : String(err));
    } finally {
      this.loading.set(false);
    }
  }

  /** Render the document with the schema's own toDOM rules. */
  private renderPreview(doc: ProseMirrorNode): void {
    const host = this.previewHost()?.nativeElement;
    if (!host) return;
    host.replaceChildren(this.serializer.serializeFragment(doc.content, { document }));
  }
}
