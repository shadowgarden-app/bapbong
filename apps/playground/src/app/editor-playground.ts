import { Component, ElementRef, OnDestroy, inject, signal, viewChild } from '@angular/core';
import { NgTemplateOutlet } from '@angular/common';
import { ReplyEditorDirective } from './reply-editor.directive';
import { CommentsStore } from './comments-store';
import { DOMSerializer, Node as ProseMirrorNode } from 'prosemirror-model';
import { BapbongEditor, type CellBlock, type EditorChange, type SelectedCell } from '@shadow-garden/bapbong-editor';
import {
  cellAt,
  deleteColumn,
  deleteRow,
  deleteSelectionCommand,
  deleteTable,
  insertColumn,
  insertImage,
  insertRow,
  insertTable,
  mergeCells,
  setLink,
} from '@shadow-garden/bapbong-commands';
import type { BorderSide, Command, EditorPointerEvent } from '@shadow-garden/bapbong-contracts';
import {
  createFindDialog,
  Dialog,
  mountMenubar,
  mountToolbar,
  openCellProperties,
  promptDialog,
  showContextMenu,
  tableGridPicker,
  type BorderPreset,
  type ContextMenuEntry,
  type FindDialogHandle,
  type Menu,
  type MenubarHandle,
  type ToolbarHandle,
} from '@shadow-garden/bapbong-ui';

/** The JSON / DOM-preview panels are inspection aids — sync them lazily. */
const PANEL_SYNC_MS = 250;

/** Resolve a border preset to one cell's four sides given its position in the
 *  block — Outside/Inside depend on which edge the cell sits on. */
function borderSidesFor(
  preset: BorderPreset,
  cell: SelectedCell,
  block: CellBlock,
  on: BorderSide,
): { top: BorderSide | false; right: BorderSide | false; bottom: BorderSide | false; left: BorderSide | false } {
  const off = false as const;
  const topRow = cell.row === 0;
  const bottomRow = cell.row === block.rows - 1;
  const leftCol = cell.col === 0;
  const rightCol = cell.col === block.cols - 1;
  switch (preset) {
    case 'all':
      return { top: on, right: on, bottom: on, left: on };
    case 'none':
      return { top: off, right: off, bottom: off, left: off };
    case 'outside':
      return { top: topRow ? on : off, right: rightCol ? on : off, bottom: bottomRow ? on : off, left: leftCol ? on : off };
    case 'inside':
      return { top: topRow ? off : on, right: rightCol ? off : on, bottom: bottomRow ? off : on, left: leftCol ? off : on };
    case 'top':
      return { top: topRow ? on : off, right: off, bottom: off, left: off };
    case 'bottom':
      return { top: off, right: off, bottom: bottomRow ? on : off, left: off };
    case 'left':
      return { top: off, right: off, bottom: off, left: leftCol ? on : off };
    case 'right':
      return { top: off, right: rightCol ? on : off, bottom: off, left: off };
    case 'insideH':
      return { top: topRow ? off : on, right: off, bottom: bottomRow ? off : on, left: off };
    case 'insideV':
      return { top: off, right: rightCol ? off : on, bottom: off, left: leftCol ? off : on };
  }
}

/**
 * The playground is a thin shell: {@link BapbongEditor} owns the canvas
 * render/edit loop, {@link CommentsStore} owns the comment subsystem, and this
 * component just wires file load → editor + the inspection panels (rendered
 * preview, document JSON). The template binds the comment UI straight to the
 * store (`cs.*`).
 */
@Component({
  selector: 'app-editor-playground',
  templateUrl: './editor-playground.html',
  styleUrl: './editor-playground.css',
  imports: [NgTemplateOutlet, ReplyEditorDirective],
  providers: [CommentsStore],
})
export class EditorPlayground implements OnDestroy {
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
      // External plugins: comment tint + a right-click context menu (find/replace
      // and table-resize are built into the editor).
      plugins: [
        this.cs.tintPlugin,
        {
          name: 'context-menu',
          onPointer: (ev) => {
            if (ev.type !== 'contextmenu') return false;
            this.openContextMenu(ev);
            return true; // claim → the editor suppresses the native menu
          },
        },
      ],
    });
    // Shell concerns: page count + the lazy inspection panels.
    editor.onChange((c) => this.onEditorChange(c));
    // Cell-block action icon → cell-properties dialog (applied to the block).
    editor.tableSelection.onAction((block) => this.openCellPropsForBlock(block));
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
          ['bullet-list', 'ordered-list'],
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
          {
            label: 'List',
            submenu: [{ command: 'bullet-list' }, { command: 'ordered-list' }],
          },
        ],
      },
      { label: 'Help', entries: [{ label: 'Keyboard shortcuts', run: () => this.showShortcuts() }] },
    ];
  }

  /** Right-click menu: 5 edit defaults + table ops when in a cell. The selection
   *  moves to the click unless it lands inside an existing selection. */
  private openContextMenu(ev: EditorPointerEvent): void {
    const ed = this.editor;
    if (!ed) return;
    const sel = ed.state.selection;
    const insideSel = !sel.empty && ev.pos != null && ev.pos >= sel.from && ev.pos <= sel.to;
    if (ev.pos != null && !insideSel) ed.setSelection(ev.pos);

    const hasSelection = !ed.state.selection.empty;
    const entries: ContextMenuEntry[] = [
      { label: 'Cut', shortcut: '⌘X', enabled: hasSelection, run: () => ed.cut() },
      { label: 'Copy', shortcut: '⌘C', enabled: hasSelection, run: () => ed.copy() },
      { label: 'Paste', shortcut: '⌘V', run: () => void ed.paste() },
      { label: 'Paste without formatting', shortcut: '⇧⌘V', run: () => void ed.pasteText() },
      { label: 'Delete', enabled: hasSelection, run: () => this.exec(deleteSelectionCommand()) },
    ];
    const cell = cellAt(ed.state);
    if (cell) {
      // Act on the selected block if there is one, else the clicked cell (1×1).
      const block = ed.tableSelection.block() ?? { cells: [{ pos: cell.pos, row: 0, col: 0 }], rows: 1, cols: 1 };
      entries.push('separator', { label: 'Cell properties…', run: () => this.openCellPropsForBlock(block) });
      if (block.cells.length > 1) {
        entries.push({ label: 'Merge cells', run: () => this.exec(mergeCells(block.cells, block.rows, block.cols)) });
      }
      entries.push(
        { label: 'Insert row above', run: () => this.exec(insertRow(false)) },
        { label: 'Insert row below', run: () => this.exec(insertRow(true)) },
        { label: 'Insert column left', run: () => this.exec(insertColumn(false)) },
        { label: 'Insert column right', run: () => this.exec(insertColumn(true)) },
        'separator',
        { label: 'Delete row', run: () => this.exec(deleteRow()) },
        { label: 'Delete column', run: () => this.exec(deleteColumn()) },
        { label: 'Delete table', run: () => this.exec(deleteTable()) },
      );
    }
    showContextMenu(entries, { x: ev.clientX, y: ev.clientY });
  }

  /** Open the cell-properties dialog for a selected block, pre-filled from its
   *  first cell. Fill + vAlign apply uniformly; a chosen border preset is
   *  resolved per cell (Outside/Inside depend on block position) — all in one
   *  undoable transaction. */
  private openCellPropsForBlock(block: CellBlock): void {
    const ed = this.editor;
    if (!ed || block.cells.length === 0) return;
    const first = ed.state.doc.nodeAt(block.cells[0].pos);
    const va = first?.attrs['vAlign'];
    openCellProperties({
      initial: {
        background: (first?.attrs['background'] as string | null) ?? null,
        vAlign: va === 'center' || va === 'bottom' ? va : 'top',
      },
      singleCell: block.rows === 1 && block.cols === 1,
      onApply: (result) => {
        const on: BorderSide | null = result.border
          ? { width: result.border.width, style: result.border.style, color: result.border.color }
          : null;
        let tr = ed.state.tr;
        for (const cell of block.cells) {
          if (ed.state.doc.nodeAt(cell.pos)?.type.name !== 'table_cell') continue;
          tr = tr
            .setNodeAttribute(cell.pos, 'background', result.background)
            .setNodeAttribute(cell.pos, 'vAlign', result.vAlign === 'top' ? null : result.vAlign);
          if (result.border && on) {
            tr = tr.setNodeAttribute(cell.pos, 'borders', borderSidesFor(result.border.preset, cell, block, on));
          }
        }
        if (tr.docChanged) ed.dispatch(tr);
        ed.tableSelection.clear();
      },
    });
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
