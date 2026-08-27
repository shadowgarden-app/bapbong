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
  internalLinkFor,
} from '@shadow-garden/bapbong-editor';
import {
  activeFontFamily,
  activeFontSize,
  activeHighlight,
  activeListPresetId,
  activeTextColor,
  applyListPreset,
  cellAt,
  deleteColumn,
  deleteRow,
  deleteSelectionCommand,
  deleteTable,
  insertColumn,
  insertImage,
  insertRow,
  insertTable,
  insertText,
  linkInfoAt,
  listPresets,
  mergeCells,
  removeSectionBreak,
  activeCharacterFormatting,
  applyCharacterFormatting,
  currentPageConfig,
  PAPER_SIZES,
  setSectionOrientation,
  setSectionPageDimensions,
  setSectionPageNumbers,
  setSectionPaperSize,
  type PaperSize,
  setFontFamily,
  setFontSize,
  setHighlight,
  setLink,
  setTextColor,
  insertEquation,
  insertEquationNode,
} from '@shadow-garden/bapbong-commands';
import type {
  BorderSide,
  Command,
  EditorPointerEvent,
  PageConfig,
  SectionConfig,
  EqNode,
  VectorOp,
} from '@shadow-garden/bapbong-contracts';
import {
  Collection,
  KeybindingRegistry,
} from '@shadow-garden/bapbong-contracts';
import {
  IS_MAC,
  installWindowKeymap,
} from '@shadow-garden/bapbong-input-bridge';
import {
  createCanvasMeasurer,
  createCanvasMetrics,
  createFontRegistryMeasurer,
  createFontRegistryMetrics,
  type FontRegistry,
} from '@shadow-garden/bapbong-measuring';
import { audit, type AuditEntry } from '@shadow-garden/bapbong-docx';
import { layoutEquation } from '@shadow-garden/bapbong-layout-engine';
import { loadBundledFonts } from './fonts';
import {
  createFindDialog,
  createSectionChip,
  createSymbolDialog,
  openKeyboardShortcutsDialog,
  openSymbolPopover,
  panelAnchor,
  type SymbolDialogHandle,
  equationGallery,
  equationPanel,
  type EquationPanel,
  mountMenubar,
  mountToolbar,
  openCellProperties,
  openFontDialog,
  openPageSizeDialog,
  pageNumberPicker,
  promptDialog,
  sectionPaperPanel,
  showContextMenu,
  showLinkPanel,
  showPopup,
  tableGridPicker,
  type BorderPreset,
  type LinkPanelHandle,
  type ContextMenuEntry,
  type SectionChipHandle,
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
/** Offered in the toolbar pickers and in Format ▸ Font, so the two cannot
 *  drift apart on what is available. */
/** localStorage key for the Symbol dialog's recently-used row. */
const RECENT_SYMBOLS_KEY = 'bapbong.recentSymbols';
function readRecentSymbols(): string[] {
  try {
    const raw = JSON.parse(localStorage.getItem(RECENT_SYMBOLS_KEY) ?? '[]');
    return Array.isArray(raw) ? raw.filter((c) => typeof c === 'string') : [];
  } catch {
    return [];
  }
}

const FONT_FAMILIES = [
  'Arial',
  'Times New Roman',
  'Georgia',
  'Calibri',
  'Courier New',
  'Verdana',
  'Tahoma',
] as const;
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48] as const;

/** One .docx the dev server found in `apps/playground/public`. */
interface SampleFile {
  name: string;
  size: number;
}

/** Shown when `/api/samples` isn't served — a built playground has no dev
 *  server behind it, so fall back to the names that have always been there. */
const BUILTIN_SAMPLES: readonly SampleFile[] = [
  { name: 'sample.docx', size: 0 },
  { name: 'large_sample.docx', size: 0 },
  { name: 'khtn6.docx', size: 0 },
  { name: 'tab-stops.docx', size: 0 },
];

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
  /** The .docx files sitting in `public/`, as reported by the dev-server API
   *  (`proxy.config.mjs`) — the sample dropdown's options. */
  protected readonly samples = signal<readonly SampleFile[]>(BUILTIN_SAMPLES);
  protected readonly selectedSample = signal('');
  /** Dev mode: run the XML coverage audit on every import and show its
   *  report inline (the playground's whole point is inspecting conversion —
   *  no console needed). Persisted under the audit's own localStorage key,
   *  so it also survives into a plain `globalThis.__BAPBONG_XML_AUDIT__`. */
  protected readonly xmlAudit = signal(audit.enabled);
  protected readonly auditUnknown = signal<AuditEntry[]>([]);
  /** Unread but provably harmless (value = the spec's no-op). Kept out of the
   *  UNKNOWN count and folded away by default — visible on demand so the
   *  demotion stays auditable rather than becoming a silent ignore-list. */
  protected readonly auditInert = signal<AuditEntry[]>([]);
  /** Declared but unreachable — see AuditReport.unreferenced. */
  protected readonly auditUnref = signal<AuditEntry[]>([]);
  protected readonly auditIgnored = signal<AuditEntry[]>([]);
  protected readonly showAuditInert = signal(false);
  protected readonly showAuditUnref = signal(false);
  protected readonly showAuditIgnored = signal(false);
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
  /** Where floating panels (find, symbol) dock: the canvas viewport. The
   *  playground's chrome sits ABOVE the wrap in normal flow, so nothing to
   *  avoid — but it goes through the same helper the desktop shell uses. */
  private readonly floatAnchor = panelAnchor({
    area: () => this.wrapHost()?.nativeElement,
  });
  /**
   * App-level shortcuts — chords that work wherever focus is (⌘F): the host's
   * own registry + commands, dispatched by one document listener. The
   * editor's own chords (⌘B, Enter, …) live in `editor.keybindings`; the
   * Keyboard-shortcuts dialog lists both.
   */
  private readonly appCommands = new Collection<Command>([], {
    idProperty: 'name',
  });
  private readonly appKeys = new KeybindingRegistry(IS_MAC);
  private disposeWindowKeys: (() => void) | null = null;
  /** Insert › Symbol… — non-modal like Word's, built once per editor. */
  private symbolDialog: SymbolDialogHandle | null = null;

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
    void this.loadSampleList();
  }
  // Debounced side panels.
  private panelTimer: ReturnType<typeof setTimeout> | null = null;

  protected async onFile(event: Event): Promise<void> {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (!file) return;
    await this.load(file.name, await file.arrayBuffer());
  }

  /** Ask the dev server which samples `public/` holds. Any failure (no dev
   *  server, SPA fallback serving index.html) just keeps the built-in list. */
  private async loadSampleList(): Promise<void> {
    try {
      const res = await fetch('/api/samples');
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const body = (await res.json()) as { files?: SampleFile[] };
      if (body.files?.length) this.samples.set(body.files);
    } catch {
      // Keep BUILTIN_SAMPLES — the dropdown still works.
    }
  }

  protected async onSamplePick(event: Event): Promise<void> {
    const name = (event.target as HTMLSelectElement).value;
    if (!name) return;
    this.selectedSample.set(name);
    await this.loadSample(name);
  }

  /** "khtn6.docx · 961 KB" — the size only when the API reported one. */
  protected sampleLabel(file: SampleFile): string {
    if (!file.size) return file.name;
    const kb = file.size / 1024;
    const size =
      kb >= 1024 ? `${(kb / 1024).toFixed(1)} MB` : `${Math.round(kb)} KB`;
    return `${file.name} · ${size}`;
  }

  protected async loadSample(name = 'sample.docx'): Promise<void> {
    try {
      // Names carry spaces and parentheses — encode before fetching.
      const res = await fetch(encodeURIComponent(name));
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
      if (!editor) throw new Error('Canvas is not ready.');
      const { headerKeys, footerKeys } = await editor.loadDocx(bytes);
      this.headerKeys.set(headerKeys);
      this.footerKeys.set(footerKeys);
      this.readAuditReport();
    } catch (err) {
      this.error.set(err instanceof Error ? err.message : String(err));
    } finally {
      this.loading.set(false);
    }
  }

  /** Pull the report the import just produced (null unless the flag is on). */
  private readAuditReport(): void {
    const report = audit.enabled ? audit.lastReport : null;
    this.auditUnknown.set(report?.unknown ?? []);
    this.auditInert.set(report?.inert ?? []);
    this.auditUnref.set(report?.unreferenced ?? []);
    this.auditIgnored.set(report?.ignored ?? []);
  }

  /** Toggle the XML coverage audit. Takes effect on the NEXT import — the
   *  audit resolves its flag once per import, so the open document keeps the
   *  report it was parsed with. */
  protected toggleXmlAudit(on: boolean): void {
    audit.setEnabled(on);
    try {
      localStorage.setItem('bapbong.xmlAudit', on ? '1' : '0');
    } catch {
      /* private mode — session only */
    }
    this.xmlAudit.set(on);
    if (!on) {
      this.auditUnknown.set([]);
      this.auditInert.set([]);
      this.auditUnref.set([]);
      this.auditIgnored.set([]);
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
      // Wider than the painter's 24px default: the section-break chip lives
      // in the gap and needs breathing room at 100% zoom.
      pageGap: 32,
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
    // onPlugin, not plugin(): instances are rebuilt per document, so a
    // one-shot subscription would die at the first load.
    editor.onPlugin('table-selection', (p) =>
      p.onAction((block) => this.openCellPropsForBlock(block)),
    );
    // The equation palette. Built once (it holds the open tab and the chosen
    // symbol set) and re-handed to each document's plugin instance.
    editor.onPlugin('equation', (p) => p.usePanel(this.equationPanel().el));
    // Menubar / toolbar / find-bar: hand bapbong-ui the host elements + the
    // editor; it renders from editor.commands / editor.find and wires everything
    // itself. The menubar tree mixes registry commands with host actions (open
    // file, comment view, find, shortcuts) and a table-size widget — see
    // buildMenus().
    // The full Symbol panel (Insert › Symbol…, and "More symbols…" from the
    // toolbar's Ω picker). Recents are a playground preference the picker and
    // the panel share through localStorage.
    this.symbolDialog = createSymbolDialog({
      recent: readRecentSymbols(),
      onInsert: (ch) => {
        this.exec(insertText(ch));
      },
      onRecentChange: (r) =>
        localStorage.setItem(RECENT_SYMBOLS_KEY, JSON.stringify(r)),
      anchor: this.floatAnchor,
    });
    const menubarHost = this.editorMenubar()?.nativeElement;
    if (menubarHost)
      this.menubar = mountMenubar(menubarHost, editor, {
        menus: this.buildMenus(),
        keybindings: [this.appKeys], // ⌘F and friends live here
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
                { label: 'Font', value: '' },
                ...FONT_FAMILIES.map((f) => ({ label: f, value: f })),
              ],
              value: (s) => activeFontFamily(s) ?? '',
              onSelect: (v) => this.exec(setFontFamily(v || null)),
            },
            {
              kind: 'select',
              title: 'Font size',
              width: 66,
              options: [
                { label: 'Size', value: '' },
                ...FONT_SIZES.map((n) => ({
                  label: String(n),
                  value: String(n),
                })),
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
              clearLabel: 'Automatic',
              value: (s) => activeTextColor(s),
              onSelect: (c) => this.exec(setTextColor(c)),
            },
            {
              kind: 'color',
              title: 'Highlight',
              glyph:
                '<svg viewBox="0 0 16 16" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"><path d="M3 13h10"/><path d="M5 11l-1 1 2 0 6.5-6.5a1.5 1.5 0 0 0-2-2L4 10z"/></svg>',
              clearLabel: 'No highlight',
              value: (s) => activeHighlight(s),
              onSelect: (c) => this.exec(setHighlight(c)),
            },
            'clear-format',
            {
              // Ω opens the quick picker under the button; "More symbols…"
              // there and Insert › Symbol… open the full panel.
              kind: 'button',
              title: 'Insert symbol',
              label: 'Ω',
              onClick: (anchor) =>
                openSymbolPopover({
                  anchor,
                  recent: readRecentSymbols(),
                  onInsert: (ch) => this.exec(insertText(ch)),
                  onRecentChange: (r) =>
                    localStorage.setItem(RECENT_SYMBOLS_KEY, JSON.stringify(r)),
                  onMore: () => this.symbolDialog?.open(),
                  onClose: () => editor.focus(),
                }),
            },
          ],
          [
            {
              kind: 'split',
              name: 'bullet-list',
              options: listPresets('bullet').map((p) => ({
                value: p.id,
                rows: p.samples,
              })),
              value: (s) => activeListPresetId(s, 'bullet'),
              onSelect: (id) => this.exec(applyListPreset('bullet', id)),
            },
            {
              kind: 'split',
              name: 'ordered-list',
              options: listPresets('ordered').map((p) => ({
                value: p.id,
                rows: p.samples,
              })),
              value: (s) => activeListPresetId(s, 'ordered'),
              onSelect: (id) => this.exec(applyListPreset('ordered', id)),
            },
          ],
          ['align-left', 'align-center', 'align-right', 'align-justify'],
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
              onSelect: (v) => {
                editor.setZoom(Number(v) / 100);
                // Zoom repaints without an editor change — the section
                // markers track page coords, so reposition them here.
                this.renderSectionMarkers();
              },
            },
          ],
        ],
      });
    // Find/replace as a (non-modal) dialog opened from Edit ▸ Find and replace,
    // pinned to the canvas viewport's top-right (like Google Docs). Uses the
    // lib's English defaults.
    // ⌘F is an app chord (it opens a panel), so it is the app registry's —
    // not the find dialog's own listener.
    this.findDialog = createFindDialog(() => editor.plugin('find'), {
      anchor: this.floatAnchor,
      shortcut: false,
    });
    this.appCommands.add({
      name: 'find',
      title: 'Find and replace',
      run: (_s, dispatch) => {
        if (dispatch) this.findDialog?.open();
        return true;
      },
    });
    this.appKeys.add({
      key: 'Mod-f',
      command: 'find',
      scope: 'window',
      source: 'playground',
    });
    this.disposeWindowKeys?.();
    this.disposeWindowKeys = installWindowKeymap({
      keybindings: this.appKeys,
      commands: this.appCommands,
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
  /** Pool of marker lines (each with an × delete button) appended to the
   *  stack — CONTINUOUS breaks only (they happen mid-page). */
  private readonly sectionMarkerEls: HTMLDivElement[] = [];
  /** Pool of next-page break markers: a dashed line across the page gap plus
   *  the segmented chip (page numbering / paper menus + delete). `boundary`
   *  is rebound on every render, so the chip's callbacks stay closures over
   *  the slot, not over a stale index. */
  private readonly sectionChipSlots: {
    line: HTMLDivElement;
    chip: SectionChipHandle;
    boundary: number;
    pageIndex: number;
  }[] = [];

  /** Draw the section-break markers. A continuous break is a dashed line at
   *  its in-page boundary (with an × to delete); a next-page break IS the
   *  page boundary, so its marker — dashed line + quick-action chip — lives
   *  centered in the gap between the two pages. Markers live in the stack, so
   *  they scroll with the pages and only need repositioning when the layout
   *  changes (on every `onChange`). */
  private renderSectionMarkers(): void {
    const ed = this.editor;
    const stack = this.stackHost()?.nativeElement;
    if (!ed || !stack) return;
    const boundaries = this.showSections() ? ed.sectionBoundaries() : [];
    const lineBs = boundaries.filter((b) => !b.newPage);
    const chipBs = boundaries.filter((b) => b.newPage);

    while (this.sectionMarkerEls.length < lineBs.length) {
      const line = document.createElement('div');
      line.style.cssText =
        'position:absolute;z-index:7;border-top:1px dashed var(--bb-ui-active-fg,#378add);pointer-events:none;';
      const x = document.createElement('button');
      x.type = 'button';
      x.setAttribute('aria-label', 'Remove section break');
      x.style.cssText =
        'position:absolute;left:-9px;top:-9px;width:18px;height:18px;display:flex;align-items:center;' +
        'justify-content:center;padding:0;border:1px solid var(--bb-ui-active-fg,#378add);border-radius:50%;background:var(--bb-ui-menu-bg,#fff);' +
        'color:var(--bb-ui-active-fg,#378add);cursor:pointer;pointer-events:auto;';
      x.innerHTML =
        '<svg viewBox="0 0 16 16" width="9" height="9" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M3 3l10 10M13 3 3 13"/></svg>';
      x.addEventListener('pointerdown', (e) => e.stopPropagation());
      x.addEventListener('mousedown', (e) => e.stopPropagation());
      line.appendChild(x);
      stack.appendChild(line);
      this.sectionMarkerEls.push(line);
    }
    this.sectionMarkerEls.forEach((line, i) => {
      const rect = lineBs[i]?.rect;
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
      const index = lineBs[i].index;
      (line.querySelector('button') as HTMLButtonElement).onclick = (e) => {
        e.stopPropagation();
        this.exec(removeSectionBreak(index));
      };
    });

    while (this.sectionChipSlots.length < chipBs.length) {
      const line = document.createElement('div');
      line.style.cssText =
        'position:absolute;z-index:7;border-top:1.5px dashed var(--bb-ui-active-fg,#378add);opacity:.7;pointer-events:none;';
      stack.appendChild(line);
      const slot = {
        line,
        boundary: 0,
        pageIndex: 0,
        chip: null as unknown as SectionChipHandle,
      };
      slot.chip = createSectionChip({
        title: 'Next page',
        ariaPageNumbers: 'Section page numbering',
        ariaPaper: 'Section paper size',
        ariaDelete: 'Remove section break',
        onPageNumbers: (anchor) => this.openSectionNumbering(slot, anchor),
        onPaper: (anchor) => this.openSectionPaper(slot, anchor),
        onDelete: () => this.exec(removeSectionBreak(slot.boundary)),
      });
      stack.appendChild(slot.chip.el);
      this.sectionChipSlots.push(slot);
    }
    this.sectionChipSlots.forEach((slot, i) => {
      const b = chipBs[i];
      const rect = b?.rect;
      const tl =
        rect && ed.pageToCanvas({ pageIndex: rect.pageIndex, x: 0, y: 0 });
      const tr =
        rect &&
        ed.pageToCanvas({ pageIndex: rect.pageIndex, x: rect.width, y: 0 });
      if (!rect || !tl || !tr) {
        slot.line.style.display = 'none';
        slot.chip.el.style.display = 'none';
        return;
      }
      slot.boundary = b.index;
      slot.pageIndex = rect.pageIndex;
      // Center of the inter-page gap (scaled by zoom, like every page coord).
      const gapY = tl.y - (ed.getPageGap() * ed.getZoom()) / 2;
      slot.line.style.display = 'block';
      slot.line.style.left = `${tl.x}px`;
      slot.line.style.top = `${gapY}px`;
      slot.line.style.width = `${tr.x - tl.x}px`;
      slot.chip.el.style.display = '';
      slot.chip.el.style.left = `${(tl.x + tr.x) / 2}px`;
      slot.chip.el.style.top = `${gapY}px`;
      const shows = ed.sectionShowsPageNumbers(slot.boundary + 1);
      slot.chip.update({
        pageNumbers: shows
          ? this.sectionNumberingLabel(slot)
          : 'no page number',
        pageNumbersMuted: !shows,
        paper: this.sectionPaperLabel(slot),
      });
    });
  }

  /** The section AFTER break `boundary` — the one the chip's menus edit. */
  private chipSection(slot: { boundary: number }): SectionConfig | null {
    const sections = this.editor?.state.doc.attrs['sections'] as
      | SectionConfig[]
      | null;
    return sections?.[slot.boundary + 1] ?? null;
  }

  /** "page ii → 1" (restart) or "page 5" (continuing). */
  private sectionNumberingLabel(slot: {
    boundary: number;
    pageIndex: number;
  }): string {
    const ed = this.editor;
    if (!ed) return '';
    const labels = ed.layout?.pageLabels;
    const cur = labels?.[slot.pageIndex] ?? String(slot.pageIndex + 1);
    // "ii → 1" is a statement about a VISIBLE transition, so it needs the page
    // before the break to actually show its number. pageLabels comes from
    // w:pgNumType alone and keeps counting through a section whose numbers are
    // hidden — quoting it there would advertise a number the reader never sees.
    const restarts = this.chipSection(slot)?.pageNumbers?.start != null;
    if (!restarts || !ed.sectionShowsPageNumbers(slot.boundary))
      return `page ${cur}`;
    const prev = labels?.[slot.pageIndex - 1] ?? String(slot.pageIndex);
    return `page ${prev} → ${cur}`;
  }

  /** Effective geometry of the chip's section (its override, else the doc). */
  private chipSectionPage(slot: { boundary: number }): PageConfig {
    const ed = this.editor;
    const docPage = ed
      ? currentPageConfig(ed.state)
      : {
          width: 794,
          height: 1123,
          margin: { top: 96, right: 96, bottom: 96, left: 96 },
        };
    return this.chipSection(slot)?.page ?? docPage;
  }

  /** "A4 portrait" / "Letter landscape" / "Custom portrait". */
  private sectionPaperLabel(slot: { boundary: number }): string {
    const page = this.chipSectionPage(slot);
    const landscape = page.width > page.height;
    const [pw, ph] = landscape
      ? [page.height, page.width]
      : [page.width, page.height];
    const paper = Object.values(PAPER_SIZES).find(
      (d) => d.width === pw && d.height === ph,
    );
    return `${paper?.label ?? 'Custom'} ${landscape ? 'landscape' : 'portrait'}`;
  }

  /** The chip's page-numbering flyout (the shared pageNumberPicker). */
  private openSectionNumbering(
    slot: { boundary: number; pageIndex: number },
    anchor: HTMLElement,
  ): void {
    const target = slot.boundary + 1;
    const pn = this.chipSection(slot)?.pageNumbers;
    const r = anchor.getBoundingClientRect();
    const ed = this.editor;
    const popup = showPopup(
      pageNumberPicker({
        shown: ed?.sectionShowsPageNumbers(target) ?? true,
        onToggleShown: (show) => {
          popup.close();
          ed?.setSectionPageNumbersShown(target, show);
        },
        fmt: pn?.fmt,
        start: pn?.start,
        onPick: (next) => {
          popup.close();
          this.exec(setSectionPageNumbers(target, next));
        },
      }),
      { x: r.left, y: r.bottom + 4 },
    );
  }

  /** The chip's paper menu: ONE flat panel (orientation tiles + paper
   *  presets + custom), scoped to the section after the break. */
  private openSectionPaper(
    slot: { boundary: number },
    anchor: HTMLElement,
  ): void {
    const target = slot.boundary + 1;
    const PAPERS: PaperSize[] = ['letter', 'legal', 'executive', 'a4', 'a5'];
    const r = anchor.getBoundingClientRect();
    const page = this.chipSectionPage(slot);
    const short = Math.min(page.width, page.height);
    const long = Math.max(page.width, page.height);
    const popup = showPopup(
      sectionPaperPanel({
        page,
        items: PAPERS.map((key) => {
          const size = PAPER_SIZES[key];
          return {
            key,
            label: size.label,
            cm: size.cm,
            px: [size.width, size.height] as const,
            active: short === size.width && long === size.height,
          };
        }),
        onOrientation: (o) => {
          popup.close();
          this.exec(setSectionOrientation(target, o));
        },
        onPick: (key) => {
          popup.close();
          this.exec(setSectionPaperSize(target, key as PaperSize));
        },
        onCustom: () => {
          popup.close();
          openPageSizeDialog({
            initial: { width: page.width, height: page.height },
            onApply: ({ width, height }) =>
              this.exec(setSectionPageDimensions(target, width, height)),
          });
        },
      }),
      { x: r.left, y: r.bottom + 4 },
    );
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
  /** Typeset an equation the way the page does — the gallery previews are
   *  the engine's own display list, measured with this editor's measurer, so
   *  a preview is the drawing the document will show. */
  private eqPanel: EquationPanel | null = null;

  /** The equation palette, built on first use. It floats over the canvas —
   *  the editor positions it against the equation being edited. */
  private equationPanel(): EquationPanel {
    if (this.eqPanel) return this.eqPanel;
    this.eqPanel = equationPanel({
      layout: (ast, sizePt) => this.typesetEquation(ast, sizePt),
      onSymbol: (ch) => this.editor?.plugin('equation').insertSymbol(ch),
      onStructure: (st) => this.editor?.plugin('equation').insertStructure(st),
    });
    return this.eqPanel;
  }

  private typesetEquation(
    ast: EqNode[],
    sizePt: number,
  ): { width: number; height: number; ops: VectorOp[] } | null {
    const reg = this.fontRegistry;
    if (!reg) return null;
    const eq = layoutEquation(
      ast,
      sizePt,
      createFontRegistryMeasurer(reg, createCanvasMeasurer()),
      createFontRegistryMetrics(reg, createCanvasMetrics()),
    );
    return { width: eq.width, height: eq.height, ops: eq.ops };
  }

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
            label: 'Orientation',
            submenu: [
              { command: 'orientation-portrait', label: 'Portrait' },
              { command: 'orientation-landscape', label: 'Landscape' },
            ],
          },
          {
            label: 'Paper size',
            submenu: [
              { command: 'paper-a4', label: 'A4' },
              { command: 'paper-a5', label: 'A5' },
              { command: 'paper-a3', label: 'A3' },
              { command: 'paper-letter', label: 'Letter' },
              { command: 'paper-legal', label: 'Legal' },
            ],
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
            shortcutOf: 'find', // the app registry's ⌘F
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
            label: 'Symbol…',
            run: () => this.symbolDialog?.open(),
          },
          {
            label: 'Equation',
            widget: (close) =>
              equationGallery({
                layout: (ast, sizePt) => this.typesetEquation(ast, sizePt),
                onPick: (ast: EqNode[]) => {
                  this.exec(insertEquationNode(ast));
                  close();
                },
                onNew: () => {
                  this.exec(insertEquation());
                  close();
                },
                newShortcut: '⌥=',
              }),
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
              {
                command: 'insert-landscape-section',
                label: 'Landscape page',
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
          { label: 'Font…', run: () => this.openFontDialog() },
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
    // Inside a generated field (a table of contents), Word's menu is about
    // the FIELD, not the text under the pointer — its entries are regenerated
    // output, so "update" is the only edit that survives.
    const activeField = ed.plugin('toc').fieldAt();
    if (activeField?.field.kind === 'toc') {
      entries.push('separator', {
        label: 'Update table of contents (page numbers)',
        run: () => {
          const n = ed.plugin('toc').updatePageNumbers();
          this.error.set(
            n > 0
              ? null
              : 'Table of contents already matches the current page numbers.',
          );
          ed.focus();
        },
      });
    }
    const cell = cellAt(ed.state);
    if (cell) {
      // Act on the selected block if there is one, else the clicked cell (1×1).
      const block = ed.plugin('table-selection').block() ?? {
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
        ed.plugin('table-selection').clear();
      },
    });
  }

  /** Whether the editor has a non-empty selection (gates Cut/Copy). */
  private hasSelection(): boolean {
    return !!this.editor && !this.editor.state.selection.empty;
  }

  /** Run a parameterized command against the live editor. */
  /** Format ▸ Font — the only way in for smallCaps, double strikethrough,
   *  baseline shift, tracking and glyph scale. */
  private openFontDialog(): void {
    const ed = this.editor;
    if (!ed) return;
    openFontDialog({
      initial: activeCharacterFormatting(ed.state),
      families: FONT_FAMILIES,
      sizes: FONT_SIZES,
      onApply: (values) => this.exec(applyCharacterFormatting(values)),
    });
  }

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
      // In-document anchors (TOC entries) show their destination + a jump,
      // and field output loses edit/unlink — decided by the editor, once.
      internal: internalLinkFor(editor, info),
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
  /** Help › Keyboard shortcuts — generated from the registries: the editor's
   *  chords and the playground's app-level ones. */
  private showShortcuts(): void {
    const editor = this.editor;
    openKeyboardShortcutsDialog({
      sources: [
        ...(editor
          ? [{ keybindings: editor.keybindings, commands: editor.commands }]
          : []),
        { keybindings: this.appKeys, commands: this.appCommands },
      ],
    });
  }

  ngOnDestroy(): void {
    if (this.panelTimer != null) clearTimeout(this.panelTimer);
    this.findDialog?.destroy();
    this.symbolDialog?.destroy();
    this.symbolDialog = null;
    this.disposeWindowKeys?.();
    this.disposeWindowKeys = null;
    this.menubar?.destroy();
    this.toolbar?.destroy();
    this.editor?.destroy();
  }
}
