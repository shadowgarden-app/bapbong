import {
  Component,
  ElementRef,
  OnDestroy,
  signal,
  viewChild,
} from '@angular/core';
import { DOMSerializer, Node as ProseMirrorNode } from 'prosemirror-model';
import {
  BapbongEditor,
  type CellBlock,
  type EditorChange,
  type SelectedCell,
} from '@shadow-garden/bapbong-editor';
import {
  activeFontFamily,
  activeFontSize,
  activeHighlight,
  activeTextColor,
  cellAt,
  deleteColumn,
  deleteRow,
  deleteSelectionCommand,
  deleteTable,
  insertColumn,
  insertImage,
  insertRow,
  insertTable,
  linkInfoAt,
  mergeCells,
  removeSectionBreak,
  setFontFamily,
  setFontSize,
  setHighlight,
  setLink,
  setTextColor,
} from '@shadow-garden/bapbong-commands';
import type {
  BorderSide,
  Command,
  EditorPointerEvent,
} from '@shadow-garden/bapbong-contracts';
import {
  createCanvasMeasurer,
  createCanvasMetrics,
  createFontRegistryMeasurer,
  createFontRegistryMetrics,
  type FontRegistry,
} from '@shadow-garden/bapbong-measuring';
import { loadBundledFonts } from './fonts';
import {
  createFindDialog,
  Dialog,
  mountMenubar,
  mountToolbar,
  openCellProperties,
  promptDialog,
  showContextMenu,
  showLinkPanel,
  tableGridPicker,
  type BorderPreset,
  type LinkPanelHandle,
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
): {
  top: BorderSide | false;
  right: BorderSide | false;
  bottom: BorderSide | false;
  left: BorderSide | false;
} {
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
      return {
        top: topRow ? on : off,
        right: rightCol ? on : off,
        bottom: bottomRow ? on : off,
        left: leftCol ? on : off,
      };
    case 'inside':
      return {
        top: topRow ? off : on,
        right: rightCol ? off : on,
        bottom: bottomRow ? off : on,
        left: leftCol ? off : on,
      };
    case 'top':
      return { top: topRow ? on : off, right: off, bottom: off, left: off };
    case 'bottom':
      return { top: off, right: off, bottom: bottomRow ? on : off, left: off };
    case 'left':
      return { top: off, right: off, bottom: off, left: leftCol ? on : off };
    case 'right':
      return { top: off, right: rightCol ? on : off, bottom: off, left: off };
    case 'insideH':
      return {
        top: topRow ? off : on,
        right: off,
        bottom: bottomRow ? off : on,
        left: off,
      };
    case 'insideV':
      return {
        top: off,
        right: rightCol ? off : on,
        bottom: off,
        left: leftCol ? off : on,
      };
  }
}

/**
 * The playground is a thin shell: {@link BapbongEditor} owns the canvas
 * render/edit loop, and this component just wires file load → editor + the
 * inspection panels (rendered preview, document JSON).
 */
@Component({
  selector: 'app-editor-playground',
  templateUrl: './editor-playground.html',
  styleUrl: './editor-playground.css',
})
export class EditorPlayground implements OnDestroy {
  protected readonly fileName = signal<string | null>(null);
  protected readonly error = signal<string | null>(null);
  protected readonly json = signal<string | null>(null);
  protected readonly headerKeys = signal<string[]>([]);
  protected readonly footerKeys = signal<string[]>([]);
  protected readonly loading = signal(false);
  protected readonly pageCount = signal(0);

  private readonly previewHost =
    viewChild<ElementRef<HTMLDivElement>>('preview');
  // The editor fills this container with one <canvas> per page (virtualized).
  private readonly stackHost =
    viewChild<ElementRef<HTMLDivElement>>('canvasStack');
  // The scroll viewport the page stack lives in (handed to the editor).
  private readonly wrapHost =
    viewChild<ElementRef<HTMLDivElement>>('canvasWrap');
  // Menubar / toolbar hosts — bapbong-ui renders + wires them. The find panel
  // is a body-level dialog (no host slot), opened from Edit ▸ Find and replace.
  private readonly editorMenubar =
    viewChild<ElementRef<HTMLDivElement>>('editorMenubar');
  private readonly editorToolbar =
    viewChild<ElementRef<HTMLDivElement>>('editorToolbar');
  private menubar: MenubarHandle | null = null;
  private toolbar: ToolbarHandle | null = null;
  private findDialog: FindDialogHandle | null = null;

  /** The framework-agnostic render/edit core (lazily created on first load). */
  private editor: BapbongEditor | null = null;

  /** Bundled metric-compatible fonts for engine-independent layout, loaded
   *  eagerly at startup; awaited before the editor is created so the first
   *  layout already measures from real font metrics. */
  protected fontRegistry: FontRegistry | null = null;
  private readonly fontsReady: Promise<void>;

  constructor() {
    this.fontsReady = loadBundledFonts().then((r) => {
      this.fontRegistry = r;
    });
  }
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
      a.download =
        (this.fileName() ?? 'document').replace(/\.docx$/i, '') +
        '-export.docx';
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
      await this.fontsReady; // real metrics ready before the first layout
      const editor = this.ensureEditor();
      if (!editor) throw new Error('Canvas chưa sẵn sàng.');
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
    const reg = this.fontRegistry;
    const editor = new BapbongEditor(stack, {
      viewport: this.wrapHost()?.nativeElement,
      // Measure from bundled font metrics (engine-independent), falling back to
      // canvas for families we don't bundle. Omitted (canvas default) if the
      // bundled fonts failed to load.
      ...(reg && {
        measureText: createFontRegistryMeasurer(reg, createCanvasMeasurer()),
        measureMetrics: createFontRegistryMetrics(reg, createCanvasMetrics()),
      }),
      // External plugin: a right-click context menu (find/replace and
      // table-resize are built into the editor).
      plugins: [
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
    // Clicking into a link pops the view panel (copy / edit / unlink) —
    // Google Docs behavior; deferred a tick so the caret pick settles first.
    // A caret that MOVES but stays within the same link keeps the panel as
    // is (it anchors to the link's start, not the cursor) — no repaint.
    editor.onCaretPick(() => {
      setTimeout(() => {
        if (this.editor !== editor || !editor.state.selection.empty) return;
        if (!linkInfoAt(editor.state)?.href) return;
        // showLinkPanel dedupes by key: a caret landing in the SAME link
        // cancels the panel's deferred close and reuses it untouched.
        this.openLinkPanel();
      });
    });
    // Cell-block action icon → cell-properties dialog (applied to the block).
    editor.tableSelection.onAction((block) =>
      this.openCellPropsForBlock(block),
    );
    // Menubar / toolbar / find-bar: hand bapbong-ui the host elements + the
    // editor; it renders from editor.commands / editor.find and wires everything
    // itself. The menubar tree mixes registry commands with host actions (open
    // file, comment view, find, shortcuts) and a table-size widget — see
    // buildMenus().
    const menubarHost = this.editorMenubar()?.nativeElement;
    if (menubarHost)
      this.menubar = mountMenubar(menubarHost, editor, {
        menus: this.buildMenus(),
      });
    const toolbarHost = this.editorToolbar()?.nativeElement;
    if (toolbarHost)
      this.toolbar = mountToolbar(toolbarHost, editor, {
        groups: [
          ['undo', 'redo'],
          [
            {
              kind: 'select',
              title: 'Font',
              width: 132,
              options: [
                { label: 'Phông chữ', value: '' },
                ...[
                  'Arial',
                  'Times New Roman',
                  'Georgia',
                  'Calibri',
                  'Courier New',
                  'Verdana',
                  'Tahoma',
                ].map((f) => ({
                  label: f,
                  value: f,
                })),
              ],
              value: (s) => activeFontFamily(s) ?? '',
              onSelect: (v) => this.exec(setFontFamily(v || null)),
            },
            {
              kind: 'select',
              title: 'Font size',
              width: 66,
              options: [
                { label: 'Cỡ chữ', value: '' },
                ...[8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48].map(
                  (n) => ({ label: String(n), value: String(n) }),
                ),
              ],
              value: (s) => {
                const sz = activeFontSize(s);
                return sz != null ? String(sz) : '';
              },
              onSelect: (v) => this.exec(setFontSize(v ? Number(v) : null)),
            },
          ],
          [
            'bold',
            'italic',
            'underline',
            'strike',
            'superscript',
            'subscript',
            {
              kind: 'color',
              title: 'Text color',
              glyph: 'A',
              allowNone: true,
              swatches: [
                '#000000',
                '#5f5e5a',
                '#888780',
                '#b4b2a9',
                '#e24b4a',
                '#d85a30',
                '#ba7517',
                '#639922',
                '#1d9e75',
                '#0f6e56',
                '#378add',
                '#185fa5',
                '#534ab7',
                '#993556',
                '#d4537e',
                '#ffffff',
              ],
              value: (s) => activeTextColor(s),
              onSelect: (c) => this.exec(setTextColor(c)),
            },
            {
              kind: 'color',
              title: 'Highlight',
              glyph:
                '<svg viewBox="0 0 16 16" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"><path d="M3 13h10"/><path d="M5 11l-1 1 2 0 6.5-6.5a1.5 1.5 0 0 0-2-2L4 10z"/></svg>',
              allowNone: true,
              swatches: [
                '#fff59d',
                '#ffe082',
                '#ffcc80',
                '#ef9a9a',
                '#f48fb1',
                '#ce93d8',
                '#90caf9',
                '#a5d6a7',
                '#80deea',
                '#e6ee9c',
                '#bcaaa4',
                '#eeeeee',
              ],
              value: (s) => activeHighlight(s),
              onSelect: (c) => this.exec(setHighlight(c)),
            },
            'clear-format',
          ],
          ['align-left', 'align-center', 'align-right', 'align-justify'],
          ['bullet-list', 'ordered-list'],
          [
            {
              kind: 'select',
              title: 'Zoom',
              width: 72,
              options: [50, 75, 90, 100, 125, 150, 200].map((n) => ({
                label: `${n}%`,
                value: String(n),
              })),
              // Zoom lives on the render core, not in editor state — read it from
              // the editor (the toolbar re-reads on every change anyway).
              value: () => String(Math.round(editor.getZoom() * 100)),
              onSelect: (v) => editor.setZoom(Number(v) / 100),
            },
          ],
        ],
      });
    // Find/replace as a (non-modal) dialog opened from Edit ▸ Find and replace,
    // pinned to the canvas viewport's top-right (like Google Docs). Uses the
    // lib's English defaults.
    this.findDialog = createFindDialog(editor.find, {
      anchor: () =>
        this.wrapHost()?.nativeElement.getBoundingClientRect() ?? null,
    });
    this.editor = editor;
    return editor;
  }

  private onEditorChange(c: EditorChange): void {
    this.pageCount.set(c.pageCount);
    if (c.docChanged) {
      this.schedulePanelSync(c.state);
      // Typing/edits move the anchor — drop an open link panel rather than
      // letting it float detached from the caret.
      this.linkPanel?.close();
      this.linkPanel = null;
      this.linkPanelKey = null;
    }
    this.renderSectionMarkers();
  }

  /** Show/hide the in-document section-break markers (View toggle). */
  protected readonly showSections = signal(true);
  /** Pool of marker lines (each with an × delete button) appended to the stack. */
  private readonly sectionMarkerEls: HTMLDivElement[] = [];

  /** Draw a thin dashed line at each section boundary with an × to delete the
   *  break. Markers live in the stack, so they scroll with the pages and only
   *  need repositioning when the layout changes (on every `onChange`). */
  private renderSectionMarkers(): void {
    const ed = this.editor;
    const stack = this.stackHost()?.nativeElement;
    if (!ed || !stack) return;
    const boundaries = this.showSections() ? ed.sectionBoundaries() : [];
    while (this.sectionMarkerEls.length < boundaries.length) {
      const line = document.createElement('div');
      line.style.cssText =
        'position:absolute;z-index:7;border-top:1px dashed #378add;pointer-events:none;';
      const x = document.createElement('button');
      x.type = 'button';
      x.setAttribute('aria-label', 'Xoá section break');
      x.style.cssText =
        'position:absolute;left:-9px;top:-9px;width:18px;height:18px;display:flex;align-items:center;' +
        'justify-content:center;padding:0;border:1px solid #378add;border-radius:50%;background:#fff;' +
        'color:#378add;cursor:pointer;pointer-events:auto;';
      x.innerHTML =
        '<svg viewBox="0 0 16 16" width="9" height="9" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M3 3l10 10M13 3 3 13"/></svg>';
      x.addEventListener('pointerdown', (e) => e.stopPropagation());
      x.addEventListener('mousedown', (e) => e.stopPropagation());
      line.appendChild(x);
      stack.appendChild(line);
      this.sectionMarkerEls.push(line);
    }
    this.sectionMarkerEls.forEach((line, i) => {
      const rect = boundaries[i]?.rect;
      const tl =
        rect &&
        ed.pageToCanvas({ pageIndex: rect.pageIndex, x: rect.x, y: rect.y });
      const tr =
        rect &&
        ed.pageToCanvas({
          pageIndex: rect.pageIndex,
          x: rect.x + rect.width,
          y: rect.y,
        });
      if (!rect || !tl || !tr) {
        line.style.display = 'none';
        return;
      }
      line.style.display = 'block';
      line.style.left = `${tl.x}px`;
      line.style.top = `${tl.y - 4}px`;
      line.style.width = `${tr.x - tl.x}px`;
      const index = boundaries[i].index;
      (line.querySelector('button') as HTMLButtonElement).onclick = (e) => {
        e.stopPropagation();
        this.exec(removeSectionBreak(index));
      };
    });
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
    host.replaceChildren(
      serializer.serializeFragment(doc.content, { document }),
    );
  }

  // ── Menubar config ───────────────────────────────────────────────
  /** The full menu tree handed to bapbong-ui. Registry commands are referenced
   *  by name; everything else (open file, find, image/link/table, shortcuts)
   *  is a host action the shell owns. */
  private buildMenus(): Menu[] {
    return [
      {
        label: 'File',
        entries: [
          {
            label: 'Open…',
            run: () => this.openFilePicker('.docx', (f) => this.loadFile(f)),
          },
          {
            label: 'Download .docx',
            run: () => void this.downloadDocx(),
            isEnabled: () => this.pageCount() > 0,
          },
          'separator',
          {
            label: 'Print',
            run: () => void this.editor?.print(),
            isEnabled: () => this.pageCount() > 0,
          },
        ],
      },
      {
        label: 'Edit',
        entries: [
          { command: 'undo', label: 'Undo' },
          { command: 'redo', label: 'Redo' },
          'separator',
          {
            label: 'Cut',
            shortcut: '⌘X',
            isEnabled: () => this.hasSelection(),
            run: () => this.editor?.cut(),
          },
          {
            label: 'Copy',
            shortcut: '⌘C',
            isEnabled: () => this.hasSelection(),
            run: () => this.editor?.copy(),
          },
          {
            label: 'Paste',
            shortcut: '⌘V',
            run: () => void this.editor?.paste(),
          },
          {
            label: 'Paste without formatting',
            shortcut: '⇧⌘V',
            run: () => void this.editor?.pasteText(),
          },
          'separator',
          {
            label: 'Find and replace',
            run: () => this.findDialog?.open(),
            shortcut: '⌘F',
          },
        ],
      },
      {
        label: 'View',
        entries: [
          {
            label: 'Show section breaks',
            isActive: () => this.showSections(),
            run: () => {
              this.showSections.update((v) => !v);
              this.renderSectionMarkers();
            },
          },
        ],
      },
      {
        label: 'Insert',
        entries: [
          {
            label: 'Image',
            submenu: [
              {
                label: 'Upload…',
                run: () =>
                  this.openFilePicker('image/*', (f) =>
                    this.insertImageFile(f),
                  ),
              },
              {
                label: 'From URL…',
                run: () => void this.insertImageFromUrl(),
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
          {
            label: 'Link…',
            run: () => this.openLinkPanel(),
          },
          {
            label: 'Remove link',
            isEnabled: () => {
              const ed = this.editor;
              if (!ed) return false;
              const cmd = setLink(null);
              return (
                (cmd.isEnabled?.(ed.state) ?? true) &&
                (cmd.isActive?.(ed.state) ?? false)
              );
            },
            run: () => this.exec(setLink(null)),
          },
          {
            label: 'Break',
            submenu: [
              { command: 'page-break', label: 'Page break' },
              {
                command: 'section-break-next-page',
                label: 'Section break (next page)',
              },
              {
                command: 'section-break-continuous',
                label: 'Section break (continuous)',
              },
            ],
          },
        ],
      },
      {
        label: 'Format',
        entries: [
          {
            label: 'Heading',
            submenu: [
              { command: 'heading-1' },
              { command: 'heading-2' },
              { command: 'heading-3' },
              { command: 'heading-4' },
              { command: 'heading-5' },
              { command: 'heading-6' },
            ],
          },
          {
            label: 'Text',
            submenu: [
              { command: 'bold' },
              { command: 'italic' },
              { command: 'underline' },
              { command: 'strike' },
              { command: 'superscript' },
              { command: 'subscript' },
              { command: 'clear-format', label: 'Clear formatting' },
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
          {
            label: 'Columns',
            submenu: [
              { command: 'columns-1', label: 'One column' },
              { command: 'columns-2', label: 'Two columns' },
              { command: 'columns-3', label: 'Three columns' },
            ],
          },
        ],
      },
      {
        label: 'Help',
        entries: [
          { label: 'Keyboard shortcuts', run: () => this.showShortcuts() },
        ],
      },
    ];
  }

  /** Right-click menu: 5 edit defaults + table ops when in a cell. The selection
   *  moves to the click unless it lands inside an existing selection. */
  private openContextMenu(ev: EditorPointerEvent): void {
    const ed = this.editor;
    if (!ed) return;
    const sel = ed.state.selection;
    const insideSel =
      !sel.empty && ev.pos != null && ev.pos >= sel.from && ev.pos <= sel.to;
    if (ev.pos != null && !insideSel) ed.setSelection(ev.pos);

    const hasSelection = !ed.state.selection.empty;
    const entries: ContextMenuEntry[] = [
      {
        label: 'Cut',
        shortcut: '⌘X',
        enabled: hasSelection,
        run: () => ed.cut(),
      },
      {
        label: 'Copy',
        shortcut: '⌘C',
        enabled: hasSelection,
        run: () => ed.copy(),
      },
      { label: 'Paste', shortcut: '⌘V', run: () => void ed.paste() },
      {
        label: 'Paste without formatting',
        shortcut: '⇧⌘V',
        run: () => void ed.pasteText(),
      },
      {
        label: 'Delete',
        enabled: hasSelection,
        run: () => this.exec(deleteSelectionCommand()),
      },
    ];
    const cell = cellAt(ed.state);
    if (cell) {
      // Act on the selected block if there is one, else the clicked cell (1×1).
      const block = ed.tableSelection.block() ?? {
        cells: [{ pos: cell.pos, row: 0, col: 0 }],
        rows: 1,
        cols: 1,
      };
      entries.push('separator', {
        label: 'Cell properties…',
        run: () => this.openCellPropsForBlock(block),
      });
      if (block.cells.length > 1) {
        entries.push({
          label: 'Merge cells',
          run: () => this.exec(mergeCells(block.cells, block.rows, block.cols)),
        });
      }
      entries.push(
        { label: 'Insert row above', run: () => this.exec(insertRow(false)) },
        { label: 'Insert row below', run: () => this.exec(insertRow(true)) },
        {
          label: 'Insert column left',
          run: () => this.exec(insertColumn(false)),
        },
        {
          label: 'Insert column right',
          run: () => this.exec(insertColumn(true)),
        },
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
          ? {
              width: result.border.width,
              style: result.border.style,
              color: result.border.color,
            }
          : null;
        let tr = ed.state.tr;
        for (const cell of block.cells) {
          if (ed.state.doc.nodeAt(cell.pos)?.type.name !== 'table_cell')
            continue;
          tr = tr
            .setNodeAttribute(cell.pos, 'background', result.background)
            .setNodeAttribute(
              cell.pos,
              'vAlign',
              result.vAlign === 'top' ? null : result.vAlign,
            );
          if (result.border && on) {
            tr = tr.setNodeAttribute(
              cell.pos,
              'borders',
              borderSidesFor(result.border.preset, cell, block, on),
            );
          }
        }
        if (tr.docChanged) ed.dispatch(tr);
        ed.tableSelection.clear();
      },
    });
  }

  /** Whether the editor has a non-empty selection (gates Cut/Copy). */
  private hasSelection(): boolean {
    return !!this.editor && !this.editor.state.selection.empty;
  }

  /** Run a parameterized command against the live editor. */
  private exec(cmd: Command): void {
    if (!this.editor) return;
    cmd.run(this.editor.state, (tr) => this.editor?.dispatch(tr));
    this.editor.focus();
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
    void this.editor?.insertImageBlob(file);
  }

  /** The open link panel + the link range it was opened for (in-link caret
   *  moves skip the reopen); closed on doc changes. */
  private linkPanel: LinkPanelHandle | null = null;
  private linkPanelKey: string | null = null;

  /** Insert > Link…: floating panel at the caret (view / edit / insert). */
  private openLinkPanel(): void {
    const editor = this.editor;
    if (!editor) return;
    const state = editor.state;
    const info = state.selection.empty ? linkInfoAt(state) : null;
    // View mode anchors at the LINK's start (stable across in-link caret
    // moves); the insert form anchors at the caret itself.
    const anchor = editor.caretViewportRect(info ? info.from : undefined);
    if (!anchor) return;
    this.linkPanelKey = info ? `${info.from}:${info.to}` : null;
    this.linkPanel = showLinkPanel({
      anchor,
      key: this.linkPanelKey ?? undefined,
      existing: info?.href ? { href: info.href, text: info.text } : null,
      hasSelection: !state.selection.empty,
      onApply: (href, text) => {
        this.exec(setLink(href, text));
        editor.focus();
      },
      onUnlink: () => {
        this.exec(setLink(null));
        editor.focus();
      },
    });
  }

  /** Insert-from-URL: fetch the bytes up front and embed them (measured, like
   *  a paste) so the image survives export. Only when the fetch fails
   *  (CORS/offline) does the raw URL go in — exported as an externally-linked
   *  picture, laid out at the 96px fallback box. */
  private async insertImageFromUrl(): Promise<void> {
    const url = await promptDialog({
      title: 'Insert image from URL',
      placeholder: 'https://…',
    });
    if (!url) return;
    const blob = await fetch(url)
      .then(async (res) => {
        if (!res.ok) return null;
        const b = await res.blob();
        return b.type.startsWith('image/') ? b : null;
      })
      .catch(() => null);
    if (blob && (await this.editor?.insertImageBlob(blob))) return;
    this.exec(insertImage(url));
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
