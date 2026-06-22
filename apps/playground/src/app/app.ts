import { Component, ElementRef, OnDestroy, inject, signal, viewChild } from '@angular/core';
import { NgTemplateOutlet } from '@angular/common';
import { ReplyEditorDirective } from './reply-editor.directive';
import { CommentsStore } from './comments-store';
import { DOMSerializer, Node as ProseMirrorNode } from 'prosemirror-model';
import { BapbongEditor, type EditorChange } from '@shadow-garden/bapbong-editor';

/** The JSON / DOM-preview panels are inspection aids — sync them lazily. */
const PANEL_SYNC_MS = 250;

/**
 * The playground is a thin shell: {@link BapbongEditor} owns the canvas
 * render/edit loop, {@link CommentsStore} owns the comment subsystem, and this
 * component just wires file load → editor + the inspection panels (rendered
 * preview, document JSON). The template binds the comment UI straight to the
 * store (`cs.*`).
 */
@Component({
  selector: 'app-root',
  templateUrl: './app.html',
  styleUrl: './app.css',
  imports: [NgTemplateOutlet, ReplyEditorDirective],
  providers: [CommentsStore],
})
export class App implements OnDestroy {
  /** Comment subsystem — state + behaviour, bound directly by the template. */
  protected readonly cs = inject(CommentsStore);

  protected readonly fileName = signal<string | null>(null);
  protected readonly error = signal<string | null>(null);
  protected readonly json = signal<string | null>(null);
  protected readonly headerKeys = signal<string[]>([]);
  protected readonly footerKeys = signal<string[]>([]);
  protected readonly loading = signal(false);
  protected readonly pageCount = signal(0);

  private readonly previewHost = viewChild<ElementRef<HTMLDivElement>>('preview');
  // The editor fills this container with one <canvas> per page (virtualized).
  private readonly stackHost = viewChild<ElementRef<HTMLDivElement>>('canvasStack');
  // The scroll viewport the page stack lives in (handed to the editor).
  private readonly wrapHost = viewChild<ElementRef<HTMLDivElement>>('canvasWrap');
  // Comment-UI hosts the store reaches into (anchored layer + composer mount).
  private readonly anchorLayer = viewChild<ElementRef<HTMLDivElement>>('anchorLayer');
  private readonly composerHost = viewChild<ElementRef<HTMLDivElement>>('composerHost');

  /** The framework-agnostic render/edit core (lazily created on first load). */
  private editor: BapbongEditor | null = null;
  // Debounced side panels.
  private panelTimer: ReturnType<typeof setTimeout> | null = null;

  protected async onFile(event: Event): Promise<void> {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (!file) return;
    await this.load(file.name, await file.arrayBuffer());
  }

  protected async loadSample(name = 'sample.docx'): Promise<void> {
    try {
      const res = await fetch(name);
      if (!res.ok) throw new Error(`${name}: HTTP ${res.status}`);
      await this.load(name, await res.arrayBuffer());
    } catch (err) {
      this.error.set(err instanceof Error ? err.message : String(err));
    }
  }

  /** Export the (edited) document back to a .docx and download it. */
  protected async downloadDocx(): Promise<void> {
    if (!this.editor) return;
    try {
      const bytes = await this.editor.exportDocx();
      const blob = new Blob([bytes as BlobPart], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = (this.fileName() ?? 'document').replace(/\.docx$/i, '') + '-export.docx';
      a.click();
      URL.revokeObjectURL(url);
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
      const editor = this.ensureEditor();
      if (!editor) throw new Error('Canvas chưa sẵn sàng.');
      this.cs.closeComposer(); // a stale composer would point at the old doc
      const { headerKeys, footerKeys } = await editor.loadDocx(bytes);
      this.headerKeys.set(headerKeys);
      this.footerKeys.set(footerKeys);
    } catch (err) {
      this.error.set(err instanceof Error ? err.message : String(err));
    } finally {
      this.loading.set(false);
    }
  }

  /** Lazily build the editor over the canvas host and wire it to the store +
   *  the shell's inspection panels. */
  private ensureEditor(): BapbongEditor | null {
    if (this.editor) return this.editor;
    const stack = this.stackHost()?.nativeElement;
    if (!stack) return null;
    const editor = new BapbongEditor(stack, {
      viewport: this.wrapHost()?.nativeElement,
      plugins: [this.cs.tintPlugin], // comment tint via the decoration pipeline
    });
    // Shell concerns: page count + the lazy inspection panels.
    editor.onChange((c) => this.onEditorChange(c));
    // Comment subsystem owns the rest (threads, anchors, tint, caret picks).
    this.cs.attachViews({
      anchorLayer: () => this.anchorLayer()?.nativeElement ?? null,
      composerHost: () => this.composerHost()?.nativeElement ?? null,
    });
    this.cs.attach(editor);
    this.editor = editor;
    return editor;
  }

  private onEditorChange(c: EditorChange): void {
    this.pageCount.set(c.pageCount);
    if (c.docChanged) this.schedulePanelSync(c.state);
  }

  /** Debounced sync of the JSON / DOM-preview inspection panels. */
  private schedulePanelSync(state: EditorChange['state']): void {
    if (this.panelTimer != null) clearTimeout(this.panelTimer);
    this.panelTimer = setTimeout(() => {
      this.panelTimer = null;
      this.json.set(JSON.stringify(state.doc.toJSON(), null, 2));
      this.renderPreview(state.doc);
    }, PANEL_SYNC_MS);
  }

  /** Render the document with its own schema's toDOM rules (the doc carries the
   *  composed schema, so plugin marks like `comment` serialize correctly). */
  private renderPreview(doc: ProseMirrorNode): void {
    const host = this.previewHost()?.nativeElement;
    if (!host) return;
    const serializer = DOMSerializer.fromSchema(doc.type.schema);
    host.replaceChildren(serializer.serializeFragment(doc.content, { document }));
  }

  ngOnDestroy(): void {
    if (this.panelTimer != null) clearTimeout(this.panelTimer);
    this.editor?.destroy();
  }
}
