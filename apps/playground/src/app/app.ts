import { Component, ElementRef, OnDestroy, inject, signal, viewChild } from '@angular/core';
import { NgTemplateOutlet } from '@angular/common';
import { ReplyEditorDirective } from './reply-editor.directive';
import { CommentsStore } from './comments-store';
import { DOMSerializer, Node as ProseMirrorNode } from 'prosemirror-model';
import { BapbongEditor, type EditorChange } from '@shadow-garden/bapbong-editor';
import { insertImage, insertTable, setLink } from '@shadow-garden/bapbong-commands';
import type { Command } from '@shadow-garden/bapbong-contracts';
import {
  createFindDialog,
  Dialog,
  mountMenubar,
  mountToolbar,
  promptDialog,
  tableGridPicker,
  type FindDialogHandle,
  type Menu,
  type MenubarHandle,
  type ToolbarHandle,
} from '@shadow-garden/bapbong-ui';

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
  // Menubar / toolbar hosts — bapbong-ui renders + wires them. The find panel
  // is a body-level dialog (no host slot), opened from Edit ▸ Find and replace.
  private readonly editorMenubar = viewChild<ElementRef<HTMLDivElement>>('editorMenubar');
  private readonly editorToolbar = viewChild<ElementRef<HTMLDivElement>>('editorToolbar');
  private menubar: MenubarHandle | null = null;
  private toolbar: ToolbarHandle | null = null;
  private findDialog: FindDialogHandle | null = null;

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
      // External plugin: comment tint. (Find/replace is built into the editor.)
      plugins: [this.cs.tintPlugin],
    });
    // Shell concerns: page count + the lazy inspection panels.
    editor.onChange((c) => this.onEditorChange(c));
    // Menubar / toolbar / find-bar: hand bapbong-ui the host elements + the
    // editor; it renders from editor.commands / editor.find and wires everything
    // itself. The menubar tree mixes registry commands with host actions (open
    // file, comment view, find, shortcuts) and a table-size widget — see
    // buildMenus().
    const menubarHost = this.editorMenubar()?.nativeElement;
    if (menubarHost) this.menubar = mountMenubar(menubarHost, editor, { menus: this.buildMenus() });
    const toolbarHost = this.editorToolbar()?.nativeElement;
    if (toolbarHost)
      this.toolbar = mountToolbar(toolbarHost, editor, {
        groups: [
          ['bold', 'italic', 'underline', 'strike'],
          ['align-left', 'align-center', 'align-right', 'align-justify'],
        ],
      });
    // Find/replace as a (non-modal) dialog opened from Edit ▸ Find and replace,
    // pinned to the canvas viewport's top-right (like Google Docs). Uses the
    // lib's English defaults.
    this.findDialog = createFindDialog(editor.find, {
      anchor: () => this.wrapHost()?.nativeElement.getBoundingClientRect() ?? null,
    });
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

  // ── Menubar config ───────────────────────────────────────────────
  /** The full menu tree handed to bapbong-ui. Registry commands are referenced
   *  by name; everything else (open file, comment view, find, image/link/table,
   *  shortcuts) is a host action the shell owns. */
  private buildMenus(): Menu[] {
    const cs = this.cs;
    const commentView = (label: string, mode: 'hide' | 'minimize' | 'expand' | 'panel') => ({
      label,
      run: () => cs.setCommentView(mode),
      isActive: () => cs.commentView() === mode,
    });
    return [
      { label: 'File', entries: [{ label: 'Open…', run: () => this.openFilePicker('.docx', (f) => this.loadFile(f)) }] },
      {
        label: 'Edit',
        entries: [
          { command: 'undo', label: 'Undo' },
          { command: 'redo', label: 'Redo' },
          'separator',
          { label: 'Find and replace', run: () => this.findDialog?.open(), shortcut: '⌘F' },
        ],
      },
      {
        label: 'View',
        entries: [
          {
            label: 'Comments',
            submenu: [
              commentView('Hide comments', 'hide'),
              commentView('Minimize comments', 'minimize'),
              commentView('Expand comments', 'expand'),
              commentView('Show all comments', 'panel'),
            ],
          },
        ],
      },
      {
        label: 'Insert',
        entries: [
          {
            label: 'Image',
            submenu: [
              { label: 'Upload…', run: () => this.openFilePicker('image/*', (f) => this.insertImageFile(f)) },
              {
                label: 'From URL…',
                run: () => this.execPrompt('Insert image from URL', 'https://…', (url) => insertImage(url)),
              },
            ],
          },
          {
            label: 'Table',
            widget: (close) =>
              tableGridPicker({
                onPick: (rows, cols) => {
                  this.exec(insertTable(rows, cols));
                  close();
                },
              }),
          },
          { label: 'Link…', run: () => this.execPrompt('Insert link', 'https://…', (href) => setLink(href)) },
          { label: 'Break', submenu: [{ command: 'page-break', label: 'Page break' }] },
        ],
      },
      {
        label: 'Format',
        entries: [
          {
            label: 'Text',
            submenu: [
              { command: 'bold' },
              { command: 'italic' },
              { command: 'underline' },
              { command: 'strike' },
              { command: 'superscript' },
              { command: 'subscript' },
            ],
          },
          {
            label: 'Align',
            submenu: [
              { command: 'align-left' },
              { command: 'align-center' },
              { command: 'align-right' },
              { command: 'align-justify' },
            ],
          },
        ],
      },
      { label: 'Help', entries: [{ label: 'Keyboard shortcuts', run: () => this.showShortcuts() }] },
    ];
  }

  /** Run a parameterized command against the live editor. */
  private exec(cmd: Command): void {
    if (!this.editor) return;
    cmd.run(this.editor.state, (tr) => this.editor?.dispatch(tr));
    this.editor.focus();
  }

  /** Ask for a value via a bapbong-ui dialog, then run the command it builds. */
  private async execPrompt(title: string, placeholder: string, build: (value: string) => Command): Promise<void> {
    const value = await promptDialog({ title, placeholder });
    if (value) this.exec(build(value));
  }

  /** Open a transient file picker and hand the chosen file to `onFile`. */
  private openFilePicker(accept: string, onFile: (file: File) => void): void {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = accept;
    input.addEventListener('change', () => {
      const file = input.files?.[0];
      if (file) onFile(file);
    });
    input.click();
  }

  private async loadFile(file: File): Promise<void> {
    await this.load(file.name, await file.arrayBuffer());
  }

  private insertImageFile(file: File): void {
    const reader = new FileReader();
    reader.onload = () => this.exec(insertImage(String(reader.result)));
    reader.readAsDataURL(file);
  }

  /** A keyboard-shortcuts list shown in a bapbong-ui Dialog. */
  private showShortcuts(): void {
    const content = document.createElement('div');
    content.style.cssText = 'font:13px system-ui;line-height:1.7';
    content.innerHTML =
      '<ul style="margin:0;padding-left:18px">' +
      '<li><b>⌘Z</b> / <b>⇧⌘Z</b> — Undo / Redo</li>' +
      '<li>Type to edit · arrows + ⇧ to select</li>' +
      '<li><b>⌘C</b> / <b>⌘V</b> — copy / paste</li>' +
      '<li>Find &amp; replace from Edit ▸ Find and replace</li>' +
      '</ul>';
    const dialog = new Dialog({ title: 'Keyboard shortcuts', modal: true });
    dialog.setContent(content);
    dialog.onClose(() => dialog.destroy());
    dialog.open();
  }

  ngOnDestroy(): void {
    if (this.panelTimer != null) clearTimeout(this.panelTimer);
    this.findDialog?.destroy();
    this.menubar?.destroy();
    this.toolbar?.destroy();
    this.editor?.destroy();
  }
}
