/**
 * The shared DocumentSession semantics over ANY ProseMirror editor state —
 * anchoring (atom-safe text hits → PM positions), unique-match rules,
 * optimistic locking, mutations as transactions.
 *
 * Hosts provide the state and the dispatch:
 *  - HeadlessSession owns an EditorState (server / tests);
 *  - the desktop WebView wraps its live editor (state + dispatch) so AI edits
 *    ride the normal transaction pipeline (undo, autosave, journal).
 * Keeping ALL semantics here means both hosts behave identically — tested
 * once, headlessly.
 */
import type { EditorState, Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import {
  AnchorError,
  ContentError,
  VersionConflictError,
  type DocBlock,
  type DocChrome,
  type DocFootnote,
  type DocImage,
  type DocImageKind,
  type DocSnapshot,
  type DocumentSession,
  type FindMatch,
  type FormatTarget,
  type Formatting,
  type ImageBox,
  type ImageBytes,
  type ImageChanges,
  type ImagePlacement,
  type ImageSource,
  type ImageSourceKind,
  type InsertAnchor,
  type MutationOptions,
  type MutationResult,
  type ChromeEdit,
  type DocStyle,
  type TocOptions,
  type TocUpdateOptions,
  type PageSetup,
  type PageSetupChange,
  type SessionCapabilities,
} from './contract.js';
import { dataUrl, sniffImage } from './image-bytes.js';
import {
  buildContent,
  columnWidths,
  indentAttr,
  lengthToPx,
  linkTarget,
  listKindOf,
  mintList,
  referencedTableStyles,
  rowNode,
  spacingAttr,
  tableChrome,
  tableGrid,
  LIST_LEVEL_INDENT,
  headingStyleLevel,
  isBuiltInStyle,
  STYLE_ATTRS,
  type Content,
  type ResolvedStyle,
  type ListKind,
  type NumberingDefs,
  type TableEdit,
  type TableStyleSource,
} from './blocks.js';

/** What a PmDocSession needs from its host. */
export interface PmSessionHost {
  getState(): EditorState;
  /** Dispatch a transaction (live editor) / apply it (headless state). */
  apply(tr: Transaction): void;
  /** Opaque, changes on every doc change — INCLUDING user edits. */
  getVersion(): string;
  meta(): { name?: string; dirty?: boolean };
  save(): Promise<void>;
  /** Current user selection, when the host has one (desktop editor). */
  selection?(): { from: number; to: number } | null;
  /** The style a new table is born with ("Table Grid"), when the host can
   *  resolve it; absent → a direct 1px grid. */
  tableStyle?(): TableStyleSource | undefined;
  /** Fetch a picture's bytes: read the file, download the attachment,
   *  rasterize the markup — whatever its `imageSources` promise. This tier
   *  never touches a disk or a network itself, and runs headless in Bun, so
   *  an SVG has to come back already rasterized. */
  readImage?(source: ImageSource): Promise<ImageBytes>;
  /** Which sources `readImage` accepts. Absent → none, and the image tools
   *  are not offered. */
  imageSources?: readonly ImageSourceKind[];
  /** A transaction for a page-setup change (the editor's own Layout
   *  commands, composed — `pageSetupTransaction` in the commands package),
   *  or null when it is already in effect; throws a sentence when it cannot
   *  be done. This tier carries no commands of its own. Absent → page_setup
   *  is not offered. */
  pageSetup?(state: EditorState, change: PageSetupChange): Transaction | null;
  /** Every footnote body the document shows, by its reference's num (the
   *  model's effectiveFootnotes). Absent → only bodies written in bapbong
   *  are read. */
  footnotes?(): Record<number, PMNode>;
  /** Lay the document out and fill in the page numbers of the table of
   *  contents at `pos` (the editor's own "Update table of contents"): how
   *  many entries changed, or null when it could not lay the document out.
   *  Absent → TOC page numbers are left pending (Word fills them). */
  refreshToc?(pos: number): Promise<number | null>;
  /** The styles the document's file defines (docx `documentStyles`).
   *  Absent → only Title / Subtitle / Normal, and no list_styles. */
  styles?(): Promise<
    {
      id: string;
      name: string;
      type: string;
      basedOn?: string;
      custom: boolean;
      isDefault: boolean;
      hidden: boolean;
    }[]
  >;
  /** What paragraph style `id` gives a paragraph (docx `styleFormatting`). */
  styleFormatting?(id: string): Promise<{
    attrs: Record<string, unknown>;
    marks: { type: string; attrs?: Record<string, unknown> }[];
  } | null>;
  /** Every section's header and footer stories as the pages show them (one
   *  entry when the sections share them). Absent → get_document reports no
   *  chrome. */
  chrome?(): {
    headers: Record<string, PMNode>;
    footers: Record<string, PMNode>;
  }[];
}

/** A4 with 1in margins — what the layout shows when the doc has no page
 *  attr, so "100%" and equal columns must mean the same here. */
const DEFAULT_CONTENT_WIDTH = 794 - 96 * 2;

interface TextBlock {
  node: PMNode;
  pos: number;
  table?: { index: number; row: number; cell: number };
}

/** A text hit resolved to absolute PM positions. */
interface Hit {
  from: number;
  to: number;
  blockIndex: number;
  context: string;
}

/** The marks that are character formatting — what "clear formatting"
 *  strips. Links, comments, footnotes, equations and carried run
 *  properties are content, not formatting, and stay. */
const CHARACTER_MARKS = [
  'strong',
  'em',
  'underline',
  'strike',
  'dstrike',
  'smallCaps',
  'textColor',
  'fontSize',
  'vertAlign',
  'letterSpacing',
  'kern',
  'charScale',
  'position',
  'highlight',
  'fontFamily',
] as const;

const MARK_BY_FLAG = {
  bold: 'strong',
  italic: 'em',
  underline: 'underline',
  strike: 'strike',
} as const;

export class PmDocSession implements DocumentSession {
  readonly capabilities: SessionCapabilities;

  constructor(private readonly host: PmSessionHost) {
    this.capabilities = {
      selection: typeof host.selection === 'function',
      images:
        typeof host.readImage === 'function' ? (host.imageSources ?? []) : [],
      pageSetup: typeof host.pageSetup === 'function',
      headerFooter: typeof host.chrome === 'function',
      styles:
        typeof host.styles === 'function' &&
        typeof host.styleFormatting === 'function',
    };
  }

  // ── reads ────────────────────────────────────────────────────────────

  async snapshot(): Promise<DocSnapshot> {
    const defs = this.numbering();
    const sectionOf = this.sectionOfTopLevel();
    const doc = this.host.getState().doc;
    const blocks: DocBlock[] = this.textblocks().map(
      ({ node, pos, table }, index) => {
        const block: DocBlock = {
          index,
          type: blockType(node),
          text: node.textContent,
        };
        if (table) block.table = table;
        if (sectionOf) block.section = sectionOf(doc.resolve(pos).index(0));
        const list = listOf(node);
        if (list) {
          block.list = {
            kind: listKindOf(list.numId, list.level, defs),
            level: list.level + 1,
          };
        }
        const images = blockImages(node).map(({ node: img }, i) => {
          const image: DocImage = {
            index: i,
            alt: String(img.attrs['alt'] ?? ''),
            width: Number(img.attrs['width']) || 0,
            height: Number(img.attrs['height']) || 0,
            rotation: Number(img.attrs['rotation']) || 0,
            kind: imageKind(img),
          };
          if (img.attrs['float']) image.float = true;
          return image;
        });
        if (images.length > 0) block.images = images;
        const links = blockLinks(node);
        if (links.length > 0) block.links = links;
        return block;
      },
    );
    const chrome = this.chrome();
    const footnotes = this.footnoteList();
    return {
      docVersion: this.host.getVersion(),
      blocks,
      ...(chrome.length > 0 ? { chrome } : {}),
      ...(footnotes.length > 0 ? { footnotes } : {}),
      meta: this.host.meta(),
    };
  }

  /** The footnotes in reference order: the number shown, and the text. */
  private footnoteList(): DocFootnote[] {
    const doc = this.host.getState().doc;
    const bodies: Record<number, PMNode> = {
      ...(this.host.footnotes?.() ?? {}),
    };
    const own =
      (doc.attrs['footnoteBodies'] as Record<string, unknown> | null) ?? {};
    for (const [key, json] of Object.entries(own)) {
      if (bodies[Number(key)] || json == null) continue;
      try {
        bodies[Number(key)] = doc.type.schema.nodeFromJSON(json);
      } catch {
        /* malformed attr data: not listed */
      }
    }
    const out: DocFootnote[] = [];
    for (const { num, shown } of footnoteRefs(doc)) {
      const body = bodies[num];
      if (!body) continue;
      let text = body.textBetween(0, body.content.size, '\n').trim();
      if (text.startsWith(shown)) text = text.slice(shown.length).trimStart();
      out.push({ number: shown, text });
    }
    return out;
  }

  /** The headers and footers with text in them, each story once with the
   *  sections that show it. */
  private chrome(): DocChrome[] {
    const perSection = this.host.chrome?.() ?? [];
    const out: DocChrome[] = [];
    const seen = new Map<string, DocChrome>();
    perSection.forEach((stories, i) => {
      for (const [part, set] of [
        ['header', stories.headers],
        ['footer', stories.footers],
      ] as const) {
        for (const [variant, story] of Object.entries(set)) {
          const text = story
            .textBetween(0, story.content.size, '\n', fieldText)
            .trim();
          if (!text) continue;
          const key = `${part}\u0000${variant}\u0000${text}`;
          const known = seen.get(key);
          if (known) known.sections.push(i + 1);
          else {
            const entry: DocChrome = { part, variant, sections: [i + 1], text };
            seen.set(key, entry);
            out.push(entry);
          }
        }
      }
    });
    return out;
  }

  async find(query: string): Promise<FindMatch[]> {
    return this.hits(query).map((h, i) => ({
      blockIndex: h.blockIndex,
      occurrence: i + 1,
      context: h.context,
    }));
  }

  async getSelection(): Promise<{ text: string; blockIndex: number } | null> {
    const sel = this.host.selection?.();
    if (!sel || sel.from === sel.to) return null;
    const state = this.host.getState();
    const text = state.doc.textBetween(sel.from, sel.to, '\n');
    const blockIndex = this.textblocks().findIndex(
      ({ node, pos }) => sel.from >= pos && sel.from <= pos + node.nodeSize,
    );
    return { text, blockIndex };
  }

  // ── mutations ────────────────────────────────────────────────────────

  async replaceText(
    oldText: string,
    newText: string,
    opts: MutationOptions = {},
  ): Promise<MutationResult> {
    this.checkVersion(opts.expectedVersion);
    const hit = this.uniqueHit(oldText, opts.occurrence);
    this.host.apply(
      this.host.getState().tr.insertText(newText, hit.from, hit.to),
    );
    return {
      docVersion: this.host.getVersion(),
      range: { from: hit.from, to: hit.from + newText.length },
    };
  }

  async insertContent(
    content: Content,
    anchor: InsertAnchor,
    opts: MutationOptions = {},
  ): Promise<MutationResult> {
    this.checkVersion(opts.expectedVersion);
    const state = this.host.getState();
    const { schema } = state;
    const tableStyle = this.host.tableStyle?.();
    const insertAt = this.anchorPos(anchor);
    const defs = this.numbering();
    const paragraphStyles = await this.stylesNamedIn(content);
    const { nodes: paragraphs, numbering } = buildContent(content, schema, {
      contentWidth: this.contentWidth(),
      tableStyle,
      numbering: defs,
      continueList: this.listBefore(state.doc.resolve(insertAt).nodeBefore),
      paragraphStyles,
    });

    const inserted = paragraphs.reduce((size, node) => size + node.nodeSize, 0);
    let tr = state.tr.insert(insertAt, paragraphs);
    if (numbering) tr = tr.setDocAttribute('numbering', numbering);
    // A table born with a style needs its definition in the document's
    // sheet for the layout to paint it (the same move insertTable makes).
    if (tableStyle?.style && schema.nodes['doc'].spec.attrs?.['tableStyles']) {
      const sheet = (state.doc.attrs['tableStyles'] ?? {}) as Record<
        string,
        unknown
      >;
      for (const id of referencedTableStyles(paragraphs)) {
        if (id === tableStyle.styleId && !sheet[id]) {
          tr = tr.setDocAttribute('tableStyles', {
            ...sheet,
            [id]: tableStyle.style,
          });
        }
      }
    }
    this.host.apply(
      keepSections(state.doc, tr, anchor.position === 'before' ? -1 : 1),
    );
    return {
      docVersion: this.host.getVersion(),
      range: { from: insertAt, to: insertAt + inserted },
    };
  }

  async applyFormatting(
    target: FormatTarget,
    format: Formatting,
    opts: MutationOptions = {},
  ): Promise<MutationResult> {
    this.checkVersion(opts.expectedVersion);
    if (typeof target === 'object' && 'fromBlock' in target) {
      return this.formatBlocks(target, format);
    }
    const hit = this.formatHit(target, opts.occurrence);
    const docStyle = await this.docStyleOf(format);
    const tr = this.formatInto(
      this.host.getState().tr,
      hit,
      this.textblocks()[hit.blockIndex],
      format,
      docStyle,
    );
    if (tr.steps.length === 0) {
      // Formatting that named no supported change — a no-op success.
      return { docVersion: this.host.getVersion() };
    }
    this.host.apply(tr);
    return {
      docVersion: this.host.getVersion(),
      range: { from: hit.from, to: hit.to },
    };
  }

  /** One formatting over a run of blocks — the whole document, the body
   *  text, the headings — in ONE transaction: a single undo step, however
   *  many paragraphs it touches. List changes stay one paragraph per call:
   *  an item joins the list right above it, which this run would read from
   *  the document as it was before the call. */
  private async formatBlocks(
    target: Extract<FormatTarget, { fromBlock: number }>,
    format: Formatting,
  ): Promise<MutationResult> {
    if (format.list !== undefined || format.listLevel !== undefined) {
      throw new ContentError(
        'list and list_level go one paragraph per call (block_index), top to bottom — each item joins the list right above it.',
      );
    }
    const blocks = this.textblocks();
    const last = blocks.length - 1;
    const to = target.toBlock ?? last;
    if (target.fromBlock > last || to > last) {
      throw new AnchorError(
        `Block ${Math.max(target.fromBlock, to)} is out of range — the document has ${blocks.length} block(s), 0 to ${last}.`,
      );
    }
    if (to < target.fromBlock) {
      throw new ContentError(
        `to_block ${to} comes before from_block ${target.fromBlock}.`,
      );
    }
    const only = target.only ?? 'all';
    const picked = blocks
      .slice(target.fromBlock, to + 1)
      .filter(
        (b) =>
          only === 'all' || (only === 'headings') === isHeadingBlock(b.node),
      );
    if (picked.length === 0)
      return { docVersion: this.host.getVersion(), formatted: 0 };
    const docStyle = await this.docStyleOf(format);
    let tr = this.host.getState().tr;
    for (const block of picked) {
      const from = block.pos + 1;
      const hit: Hit = {
        from,
        to: from + block.node.content.size,
        blockIndex: blocks.indexOf(block),
        context: '',
      };
      tr = this.formatInto(tr, hit, block, format, docStyle);
    }
    if (tr.steps.length === 0) {
      return { docVersion: this.host.getVersion(), formatted: 0 };
    }
    this.host.apply(tr);
    const end = picked[picked.length - 1];
    return {
      docVersion: this.host.getVersion(),
      range: {
        from: picked[0].pos,
        to: end.pos + end.node.nodeSize,
      },
      formatted: picked.length,
    };
  }

  /** The document's own style a formatting names, resolved once. */
  private async docStyleOf(format: Formatting): Promise<ResolvedStyle | null> {
    return format.style && !isBuiltInStyle(format.style)
      ? await this.resolveStyle(format.style)
      : null;
  }

  /** Add one block's formatting to `tr`. Marks and paragraph attributes
   *  only — no step moves a position, so blocks read before the first step
   *  stay addressable after it. */
  private formatInto(
    tr: Transaction,
    hit: Hit,
    block: TextBlock,
    format: Formatting,
    docStyle: ResolvedStyle | null,
  ): Transaction {
    const { schema } = this.host.getState();
    const styleAttrs: Record<string, unknown> = {};
    if (docStyle) {
      // The style's look replaces the paragraph's: its character formatting
      // over the whole paragraph (links, comments, footnotes stay), its
      // paragraph formatting in place of what the paragraph had.
      const from = block.pos + 1;
      const to = from + block.node.content.size;
      if (to > from) {
        for (const name of CHARACTER_MARKS) {
          const type = schema.marks[name];
          if (type) tr = tr.removeMark(from, to, type);
        }
        for (const m of docStyle.marks)
          tr = tr.addMark(from, to, schema.markFromJSON(m));
      }
      for (const key of STYLE_ATTRS)
        if (key in block.node.attrs)
          styleAttrs[key] = docStyle.attrs[key] ?? null;
      const heading = docStyle.attrs['heading'] as number | null | undefined;
      styleAttrs['heading'] = heading ?? null;
      styleAttrs['styleId'] = heading ? null : docStyle.id;
    }
    if (format.clear && hit.to > hit.from) {
      for (const name of CHARACTER_MARKS) {
        const type = schema.marks[name];
        if (type) tr = tr.removeMark(hit.from, hit.to, type);
      }
    }
    for (const [flag, markName] of Object.entries(MARK_BY_FLAG)) {
      const want = format[flag as keyof typeof MARK_BY_FLAG];
      if (want === undefined) continue;
      const mark = schema.marks[markName];
      if (!mark) continue;
      tr = want
        ? tr.addMark(hit.from, hit.to, mark.create())
        : tr.removeMark(hit.from, hit.to, mark);
    }
    if (
      format.fontSize !== undefined &&
      schema.marks['fontSize'] &&
      hit.to > hit.from
    ) {
      tr = tr.addMark(
        hit.from,
        hit.to,
        schema.marks['fontSize'].create({ size: format.fontSize }),
      );
    }
    // A value mark: absent = keep, null = remove, a value = set.
    const value = <T>(
      v: T | null | undefined,
      make: (v: T) => Record<string, unknown>,
    ): Record<string, unknown> | null | undefined =>
      v === undefined ? undefined : v === null ? null : make(v);
    const valued: [string, Record<string, unknown> | null | undefined][] = [
      ['link', value(format.link, (href) => ({ href: linkTarget(href) }))],
      ['fontFamily', value(format.fontFamily, (family) => ({ family }))],
      ['textColor', value(format.color, (color) => ({ color }))],
      ['highlight', value(format.highlight, (color) => ({ color }))],
      [
        'vertAlign',
        value(format.verticalAlign, (v) => ({
          value: v === 'subscript' ? 'sub' : 'super',
        })),
      ],
    ];
    for (const [name, attrs] of valued) {
      const type = schema.marks[name];
      if (attrs === undefined || !type || hit.to <= hit.from) continue;
      tr = tr.removeMark(hit.from, hit.to, type);
      if (attrs) tr = tr.addMark(hit.from, hit.to, type.create(attrs));
    }
    const pAttrs: Record<string, unknown> = { ...styleAttrs };
    if (format.align) pAttrs['align'] = format.align;
    if (format.heading !== undefined) {
      pAttrs['heading'] = format.heading ? format.heading : null;
      if (format.heading) pAttrs['styleId'] = null;
    }
    const headingStyle = format.style ? headingStyleLevel(format.style) : null;
    if (headingStyle) {
      pAttrs['heading'] = headingStyle;
      pAttrs['styleId'] = null;
    } else if (format.style !== undefined && !docStyle) {
      // Built in: Title / Subtitle (the layout sizes them), Normal = none.
      const id = format.style === 'Normal' ? null : format.style;
      pAttrs['styleId'] = id;
      if (id) pAttrs['heading'] = null;
    }
    if (format.tabs !== undefined) {
      const width = this.contentWidth();
      pAttrs['tabs'] = format.tabs.length
        ? format.tabs.map((t) => ({
            pos: Math.round(lengthToPx(t.at, width)),
            val: t.align ?? 'left',
            ...(t.leader ? { leader: t.leader } : {}),
          }))
        : null;
    }
    if (format.list !== undefined || format.listLevel !== undefined) {
      const change = this.listChange(block.node, block.pos, format);
      Object.assign(pAttrs, change.attrs);
      if (change.numbering)
        tr = tr.setDocAttribute('numbering', change.numbering);
    }
    if (
      format.spaceBefore !== undefined ||
      format.spaceAfter !== undefined ||
      format.lineSpacing !== undefined ||
      format.indent !== undefined
    ) {
      const spacing = spacingAttr(
        block.node.attrs['spacing'] as Record<string, unknown> | null,
        format,
      );
      if (spacing !== undefined) pAttrs['spacing'] = spacing;
      // After a list change, from the indent it left (its level shift).
      const indent = indentAttr(
        ('indent' in pAttrs
          ? pAttrs['indent']
          : block.node.attrs['indent']) as Record<string, unknown> | null,
        format.indent,
        this.contentWidth(),
      );
      if (indent !== undefined) pAttrs['indent'] = indent;
    }
    if (Object.keys(pAttrs).length > 0) {
      tr = tr.setNodeMarkup(block.pos, undefined, {
        ...block.node.attrs,
        ...pAttrs,
      });
    }
    return tr;
  }

  async updateImage(
    blockIndex: number,
    imageIndex: number,
    changes: ImageChanges,
    opts: MutationOptions = {},
  ): Promise<MutationResult> {
    this.checkVersion(opts.expectedVersion);
    const img = this.imageAt(blockIndex, imageIndex);
    let tr = this.host.getState().tr;
    if (changes.width !== undefined)
      tr = tr.setNodeAttribute(
        img.pos,
        'width',
        Math.max(1, Math.round(changes.width)),
      );
    if (changes.height !== undefined)
      tr = tr.setNodeAttribute(
        img.pos,
        'height',
        Math.max(1, Math.round(changes.height)),
      );
    if (changes.rotation !== undefined) {
      tr = tr.setNodeAttribute(
        img.pos,
        'rotation',
        ((changes.rotation % 360) + 360) % 360,
      );
    }
    if (tr.steps.length === 0) return { docVersion: this.host.getVersion() };
    this.host.apply(tr);
    return {
      docVersion: this.host.getVersion(),
      range: { from: img.pos, to: img.pos + 1 },
    };
  }

  async insertImage(
    source: ImageSource,
    anchor: InsertAnchor,
    placement: ImagePlacement = {},
    opts: MutationOptions = {},
  ): Promise<MutationResult & ImageBox> {
    this.checkVersion(opts.expectedVersion);
    const picture = await this.fetchImage(source);
    const state = this.host.getState();
    const imageType = state.schema.nodes['image'];
    const paragraph = state.schema.nodes['paragraph'];
    if (!imageType || !paragraph) {
      throw new ContentError('This document cannot hold pictures.');
    }
    const box = this.boxFor(picture, placement.width);
    const node = paragraph.create(null, [
      imageType.create({
        src: dataUrl(picture.bytes, picture.mediaType),
        alt: placement.alt ?? '',
        width: box.width,
        height: box.height,
      }),
    ]);
    const at = this.anchorPos(anchor);
    this.host.apply(
      keepSections(
        state.doc,
        state.tr.insert(at, node),
        anchor.position === 'before' ? -1 : 1,
      ),
    );
    return {
      docVersion: this.host.getVersion(),
      range: { from: at, to: at + node.nodeSize },
      ...box,
    };
  }

  async replaceImage(
    blockIndex: number,
    imageIndex: number,
    source: ImageSource,
    placement: ImagePlacement = {},
    opts: MutationOptions = {},
  ): Promise<MutationResult & ImageBox> {
    this.checkVersion(opts.expectedVersion);
    const picture = await this.fetchImage(source);
    const img = this.imageAt(blockIndex, imageIndex);
    const old = img.node.attrs;
    // The box it takes over, so the page around it does not move: its width,
    // with the height following the new picture's proportions.
    const box = this.boxFor(
      picture,
      placement.width ?? (Number(old['width']) || undefined),
    );
    // Everything that described the OLD picture goes with it — above all
    // `rawDrawing`, which the exporter writes back verbatim and would keep
    // emitting the drawing this call was meant to replace.
    const node = img.node.type.create(
      {
        src: dataUrl(picture.bytes, picture.mediaType),
        alt: placement.alt ?? String(old['alt'] ?? ''),
        title: old['title'] ?? null,
        width: box.width,
        height: box.height,
        float: old['float'] ?? null,
      },
      null,
      img.node.marks,
    );
    this.host.apply(
      this.host
        .getState()
        .tr.replaceWith(img.pos, img.pos + img.node.nodeSize, node),
    );
    return {
      docVersion: this.host.getVersion(),
      range: { from: img.pos, to: img.pos + node.nodeSize },
      ...box,
    };
  }

  async deleteImage(
    blockIndex: number,
    imageIndex: number,
    opts: MutationOptions = {},
  ): Promise<MutationResult> {
    this.checkVersion(opts.expectedVersion);
    const img = this.imageAt(blockIndex, imageIndex);
    this.host.apply(
      this.host.getState().tr.delete(img.pos, img.pos + img.node.nodeSize),
    );
    return {
      docVersion: this.host.getVersion(),
      range: { from: img.pos, to: img.pos },
    };
  }

  async deleteBlocks(
    blockIndex: number,
    count: number,
    opts: MutationOptions = {},
  ): Promise<MutationResult & { deleted: number }> {
    this.checkVersion(opts.expectedVersion);
    const blocks = this.textblocks();
    if (!Number.isInteger(count) || count < 1) {
      throw new ContentError('count is how many blocks to delete: 1 or more.');
    }
    if (
      !Number.isInteger(blockIndex) ||
      blockIndex < 0 ||
      blockIndex + count > blocks.length
    ) {
      throw new AnchorError(
        `Blocks ${blockIndex}..${blockIndex + count - 1} are out of range — the document has ${blocks.length} block(s). ` +
          'Block indexes change with every edit; call get_document again.',
      );
    }
    const doomed = blocks.slice(blockIndex, blockIndex + count);
    const state = this.host.getState();

    // A table cell holds at least one paragraph: the last one cannot go.
    const perCell = new Map<number, { left: number; block: TextBlock }>();
    for (const b of doomed) {
      if (!b.table) continue;
      const $pos = state.doc.resolve(b.pos);
      const cell = $pos.before();
      const entry = perCell.get(cell) ?? {
        left: $pos.parent.childCount,
        block: b,
      };
      entry.left--;
      perCell.set(cell, entry);
    }
    for (const { left, block } of perCell.values()) {
      if (left > 0 || !block.table) continue;
      const { index, row, cell } = block.table;
      throw new ContentError(
        `Block ${blocks.indexOf(block)} is the last paragraph of its table cell (table ${index}, row ${row}, cell ${cell}), ` +
          'and a cell cannot be empty. Clear its text with replace_text, or remove the row, the column or the ' +
          'whole table with edit_table.',
      );
    }

    let tr = state.tr;
    const topLevel = doomed.filter((b) => !b.table).length;
    if (topLevel === state.doc.childCount) {
      // Everything goes: a document still has one (empty) paragraph.
      tr = tr.replaceWith(
        0,
        state.doc.content.size,
        state.schema.nodes['paragraph'].create(),
      );
    } else {
      for (const b of [...doomed].reverse()) {
        tr = tr.delete(b.pos, b.pos + b.node.nodeSize);
      }
    }
    this.host.apply(keepSections(state.doc, tr, 1));
    const at = tr.mapping.map(doomed[0].pos);
    return {
      docVersion: this.host.getVersion(),
      range: { from: at, to: at },
      deleted: count,
    };
  }

  async pageSetup(
    setup: PageSetup,
    opts: MutationOptions = {},
  ): Promise<MutationResult & { sections: number }> {
    this.checkVersion(opts.expectedVersion);
    if (!this.host.pageSetup) {
      throw new ContentError('This document cannot change its page setup.');
    }
    const state = this.host.getState();
    const change: PageSetupChange = {};
    if (setup.section !== undefined) change.section = setup.section - 1;
    if (setup.orientation) change.orientation = setup.orientation;
    if (setup.paper)
      change.paper = setup.paper.toLowerCase() as PageSetupChange['paper'];
    if (setup.columns !== undefined) change.columns = setup.columns;
    if (setup.removeSectionBreak !== undefined)
      change.removeSectionBreak = setup.removeSectionBreak - 1;
    if (setup.margins !== undefined) {
      if (typeof setup.margins === 'string') change.margins = setup.margins;
      else {
        const m: Record<string, number> = {};
        for (const side of ['top', 'right', 'bottom', 'left'] as const) {
          const v = setup.margins[side];
          if (v === undefined) continue;
          if (typeof v === 'string' && v.trim().endsWith('%')) {
            throw new ContentError(
              `A margin is a length ("2cm", "1in"), not a share of the page: ${JSON.stringify(v)}.`,
            );
          }
          m[side] = Math.round(lengthToPx(v, 0));
        }
        change.margins = m;
      }
    }
    if (setup.sectionBreakAfter) {
      const { blockIndex, newPage } = setup.sectionBreakAfter;
      const blocks = this.textblocks();
      const block = blocks[blockIndex];
      if (!block) {
        throw new AnchorError(
          `blockIndex ${blockIndex} is out of range — the document has ${blocks.length} block(s).`,
        );
      }
      if (block.table) {
        throw new ContentError(
          `Block ${blockIndex} is inside a table — a section break goes after a paragraph outside tables (or after the table: use the block right after it).`,
        );
      }
      change.sectionBreak = {
        after: state.doc.resolve(block.pos).index(0),
        newPage: newPage ?? true,
      };
    }
    let tr: Transaction | null;
    try {
      tr = this.host.pageSetup(state, change);
    } catch (err) {
      throw new ContentError(err instanceof Error ? err.message : String(err));
    }
    if (tr) this.host.apply(tr);
    const sections = this.host.getState().doc.attrs['sections'] as
      | unknown[]
      | null;
    return {
      docVersion: this.host.getVersion(),
      sections: sections?.length || 1,
    };
  }

  async editChrome(
    edit: ChromeEdit,
    opts: MutationOptions = {},
  ): Promise<MutationResult & { sections: number[] }> {
    this.checkVersion(opts.expectedVersion);
    const state = this.host.getState();
    const { schema } = state;
    if (
      !this.host.chrome ||
      !schema.nodes['doc'].spec.attrs?.['sectionChromeOverrides']
    ) {
      throw new ContentError(
        'This document cannot change its headers and footers.',
      );
    }
    if ((edit.content === undefined) === (edit.replace === undefined)) {
      throw new ContentError(
        'Pass content (to rewrite it) or old_text with new_text (to replace inside it) — one of the two.',
      );
    }
    const perSection = this.host.chrome();
    const sectionCount = Math.max(
      (state.doc.attrs['sections'] as unknown[] | null)?.length ?? 1,
      perSection.length,
      1,
    );
    if (
      edit.section !== undefined &&
      (edit.section < 1 || edit.section > sectionCount)
    ) {
      throw new ContentError(
        `No section ${edit.section} — the document has ${sectionCount} section(s).`,
      );
    }
    const targets =
      edit.section !== undefined
        ? [edit.section - 1]
        : Array.from({ length: sectionCount }, (_x, i) => i);
    const variant = edit.variant ?? 'default';
    const key = edit.part === 'header' ? 'headers' : 'footers';
    // A flat chrome is one entry that every section shows.
    const storyOf = (i: number): PMNode | undefined =>
      (perSection.length === 1 ? perSection[0] : perSection[i])?.[key]?.[
        variant
      ];
    const noun = `${variant === 'first' ? 'first-page ' : variant === 'even' ? 'even-page ' : ''}${edit.part}`;
    if (variant !== 'default') {
      for (const i of targets) {
        if (!storyOf(i)) {
          throw new ContentError(
            `Section ${i + 1} has no ${noun} — bapbong edits the ones a document already shows. Use variant "default".`,
          );
        }
      }
    }

    const changed = new Map<number, PMNode>();
    if (edit.content !== undefined) {
      const { nodes, numbering } = buildContent(edit.content, schema, {
        contentWidth: this.contentWidth(),
      });
      if (numbering) {
        throw new ContentError(
          'A header or footer cannot hold a list — write the items as plain paragraphs.',
        );
      }
      const story = schema.node(
        'doc',
        null,
        nodes.length > 0 ? nodes : [schema.nodes['paragraph'].create()],
      );
      for (const i of targets) changed.set(i, story);
    } else {
      const r = edit.replace as NonNullable<ChromeEdit['replace']>;
      let missing: Error | null = null;
      for (const i of targets) {
        const story = storyOf(i);
        const where = `the ${noun} of section ${i + 1}`;
        const hits = story ? hitsIn(textblocksOf(story), r.oldText) : [];
        // Across every section a story without the text is left alone;
        // only "found nowhere" (or an ambiguous anchor) is an error.
        if (hits.length === 0 && edit.section === undefined) {
          missing ??= new AnchorError(
            `Text not found in any ${noun}: ${JSON.stringify(r.oldText)}. get_document lists them under chrome.`,
          );
          continue;
        }
        if (!story) {
          throw new AnchorError(
            `Section ${i + 1} has no ${noun} to change — write one with content.`,
          );
        }
        const hit = pickHit(hits, r.oldText, r.occurrence, where);
        changed.set(i, replaceInStory(story, hit, r.newText));
      }
      if (changed.size === 0)
        throw missing ?? new AnchorError('Nothing to change.');
    }

    const overrides = {
      ...((state.doc.attrs['sectionChromeOverrides'] as Record<
        string,
        Record<string, Record<string, unknown>>
      > | null) ?? {}),
    };
    for (const [i, story] of changed) {
      const entry = { ...(overrides[String(i)] ?? {}) };
      entry[key] = { ...(entry[key] ?? {}), [variant]: story.toJSON() };
      overrides[String(i)] = entry;
    }
    this.host.apply(
      state.tr.setDocAttribute('sectionChromeOverrides', overrides),
    );
    return {
      docVersion: this.host.getVersion(),
      sections: [...changed.keys()].map((i) => i + 1),
    };
  }

  async insertFootnote(
    anchor: { text: string; occurrence?: number },
    content: Content,
    opts: MutationOptions = {},
  ): Promise<MutationResult & { number: string }> {
    this.checkVersion(opts.expectedVersion);
    const state = this.host.getState();
    const { schema } = state;
    const footnote = schema.marks['footnote'];
    const sup = schema.marks['vertAlign'];
    if (
      !footnote ||
      !sup ||
      !schema.nodes['doc'].spec.attrs?.['footnoteBodies']
    ) {
      throw new ContentError('This document cannot hold footnotes.');
    }
    const { nodes, numbering } = buildContent(content, schema, {
      contentWidth: this.contentWidth(),
    });
    if (numbering || nodes.some((n) => n.type.name !== 'paragraph')) {
      throw new ContentError(
        'A footnote holds paragraphs of text — no tables or lists.',
      );
    }
    let linked = false;
    for (const n of nodes)
      n.descendants((c) => {
        if (c.marks.some((m) => m.type.name === 'link')) linked = true;
        return !linked;
      });
    if (linked) {
      throw new ContentError(
        'A footnote cannot hold a link yet — write the address as text.',
      );
    }
    const hit = this.uniqueHit(anchor.text, anchor.occurrence);
    const at = hit.to;

    const refs = footnoteRefs(state.doc);
    const bodies =
      (state.doc.attrs['footnoteBodies'] as Record<string, unknown> | null) ??
      {};
    const num =
      Math.max(
        0,
        ...refs.map((r) => r.num),
        ...Object.keys(bodies).map(Number),
      ) + 1;
    // Word numbers footnotes in order; a reference shows its place.
    const numbered = (r: { shown: string }) => /^\d+$/.test(r.shown);
    const display = refs.filter((r) => r.pos < at && numbered(r)).length + 1;
    let tr = state.tr.insert(
      at,
      schema.text(String(display), [
        sup.create({ value: 'super' }),
        footnote.create({ num, id: null }),
      ]),
    );
    // The ones after it move up one.
    const later = footnoteRefs(tr.doc).filter(
      (r) => r.pos > at && r.num !== num && numbered(r),
    );
    for (const r of later.reverse()) {
      const node = tr.doc.nodeAt(r.pos);
      if (!node) continue;
      tr = tr.replaceWith(
        r.pos,
        r.pos + node.nodeSize,
        schema.text(String(Number(r.shown) + 1), node.marks),
      );
    }
    const story = schema.node(
      'doc',
      null,
      nodes.length > 0 ? nodes : [schema.nodes['paragraph'].create()],
    );
    tr = tr.setDocAttribute('footnoteBodies', {
      ...bodies,
      [String(num)]: story.toJSON(),
    });
    this.host.apply(tr);
    return {
      docVersion: this.host.getVersion(),
      range: { from: at, to: at + String(display).length },
      number: String(display),
    };
  }

  async insertToc(
    options: TocOptions,
    anchor: InsertAnchor,
    opts: MutationOptions = {},
  ): Promise<
    MutationResult & { entries: number; pageNumbers: 'updated' | 'pending' }
  > {
    this.checkVersion(opts.expectedVersion);
    const levels = options.levels ?? 3;
    if (!Number.isInteger(levels) || levels < 1 || levels > 9) {
      throw new ContentError(`levels is 1 to 9 — not ${levels}.`);
    }
    const state = this.host.getState();
    const { schema } = state;
    const link = schema.marks['link'];
    const pAttrs = schema.nodes['paragraph'].spec.attrs ?? {};
    if (!link || !('field' in pAttrs) || !('bookmarks' in pAttrs)) {
      throw new ContentError('This document cannot hold a table of contents.');
    }
    const insertAt = this.anchorPos(anchor);
    if (state.doc.resolve(insertAt).depth > 0) {
      throw new ContentError(
        'A table of contents goes between paragraphs, not inside a table — anchor it on a paragraph outside tables.',
      );
    }
    const headings = this.textblocks().filter(({ node }) => {
      const h = node.attrs['heading'] as number | null;
      return !!h && h <= levels && node.textContent.trim().length > 0;
    });
    if (headings.length === 0) {
      throw new ContentError(
        `There are no headings (levels 1-${levels}) to list. Make the section titles headings first (apply_formatting heading).`,
      );
    }

    let tr = state.tr;
    const entries = this.tocEntries(tr, headings, levels);
    const title = options.title?.trim()
      ? schema.nodes['paragraph'].create({ spacing: { after: 11 } }, [
          schema.text(options.title.trim(), [
            schema.marks['strong'].create(),
            ...(schema.marks['fontSize']
              ? [schema.marks['fontSize'].create({ size: 16 })]
              : []),
          ]),
        ])
      : null;
    const nodes = title ? [title, ...entries] : entries;
    tr = tr.insert(insertAt, nodes);
    this.host.apply(
      keepSections(state.doc, tr, anchor.position === 'before' ? -1 : 1),
    );
    const size = nodes.reduce((n, node) => n + node.nodeSize, 0);
    const firstEntry = insertAt + (title ? title.nodeSize : 0) + 1;
    const updated = this.host.refreshToc
      ? await this.host.refreshToc(firstEntry).catch(() => null)
      : null;
    return {
      docVersion: this.host.getVersion(),
      range: { from: insertAt, to: insertAt + size },
      entries: entries.length,
      pageNumbers: updated === null ? 'pending' : 'updated',
    };
  }

  async updateToc(
    options: TocUpdateOptions,
    opts: MutationOptions = {},
  ): Promise<
    MutationResult & {
      tables: number;
      entries: number;
      changed: number;
      rebuilt: boolean;
      pageNumbers: 'updated' | 'pending';
    }
  > {
    this.checkVersion(opts.expectedVersion);
    const spans = tocSpans(this.host.getState().doc);
    if (spans.length === 0) {
      throw new ContentError(
        'This document has no table of contents to update — insert_toc adds one.',
      );
    }
    const rebuilt = !!options.rebuild;
    if (rebuilt) {
      // Word's "Update entire table": the entries again from the headings
      // as they are now (added, renamed, removed), each table keeping its
      // own levels. Back to front, so earlier spans keep their positions.
      const state = this.host.getState();
      let tr = state.tr;
      for (const span of [...spans].reverse()) {
        const levels = tocLevels(span.field.instr);
        // Positions as of `tr`: a later table may already be replaced.
        const headings = this.textblocks()
          .filter(({ node }) => {
            const h = node.attrs['heading'] as number | null;
            return !!h && h <= levels && node.textContent.trim().length > 0;
          })
          .map(({ node, pos }) => ({ node, pos: tr.mapping.map(pos) }));
        if (headings.length === 0) {
          throw new ContentError(
            `There are no headings (levels 1-${levels}) left to list — the table of contents was left as it is.`,
          );
        }
        // Bookmarks first: they change attrs only, so the span's positions
        // still hold for the replacement.
        const entries = this.tocEntries(tr, headings, levels, span.field.instr);
        tr = tr.replaceWith(span.from, span.to, entries);
      }
      this.host.apply(tr);
    }
    // Page numbers from the layout, table by table (positions re-read: a
    // rebuild or an earlier refresh may have moved them).
    const refresh = this.host.refreshToc?.bind(this.host);
    let changed = 0;
    let laidOut = !!refresh;
    const count = tocSpans(this.host.getState().doc).length;
    for (let i = 0; refresh && i < count && laidOut; i++) {
      const span = tocSpans(this.host.getState().doc)[i];
      const n = span ? await refresh(span.from + 1).catch(() => null) : null;
      if (n === null) laidOut = false;
      else changed += n;
    }
    const after = tocSpans(this.host.getState().doc);
    return {
      docVersion: this.host.getVersion(),
      tables: after.length,
      entries: after.reduce((n, sp) => n + sp.count, 0),
      changed,
      rebuilt,
      pageNumbers: laidOut ? 'updated' : 'pending',
    };
  }

  /** The entries of a table of contents listing `headings`: one field
   *  shared by all of them (it is what makes them one TOC), a right tab
   *  with a dot leader to the page number, and the whole entry a link to
   *  its heading — which gets a bookmark in `tr` when it has none. The
   *  number is a placeholder until something lays the document out;
   *  `dirty` asks Word to redo it. */
  private tocEntries(
    tr: Transaction,
    headings: { node: PMNode; pos: number }[],
    levels: number,
    instr = `TOC \\o "1-${levels}" \\h \\z \\u`,
  ): PMNode[] {
    const { schema } = tr.doc.type;
    const link = schema.marks['link'];
    // Every listed heading needs a bookmark for its entry to jump to —
    // Word's own "_Toc" names, reused when a heading already has one.
    const taken = new Set<string>();
    tr.doc.descendants((n) => {
      for (const b of (n.attrs['bookmarks'] as string[] | null) ?? [])
        taken.add(b);
      return true;
    });
    const mint = () => {
      let name: string;
      do name = `_Toc${Math.floor(1e8 + Math.random() * 9e8)}`;
      while (taken.has(name));
      taken.add(name);
      return name;
    };
    const anchors = headings.map(({ node, pos }) => {
      const cur = tr.doc.nodeAt(pos) ?? node;
      const own = (cur.attrs['bookmarks'] as string[] | null) ?? [];
      const existing = own.find((b) => b.startsWith('_Toc'));
      if (existing) return existing;
      const name = mint();
      tr.setNodeAttribute(pos, 'bookmarks', [...own, name]);
      return name;
    });
    const field = { kind: 'toc', instr, dirty: true };
    const width = Math.round(this.contentWidth());
    return headings.map(({ node }, i) => {
      const level = node.attrs['heading'] as number;
      return schema.nodes['paragraph'].create(
        {
          field,
          styleId: `TOC${level}`,
          indent: level > 1 ? { left: (level - 1) * 15 } : null,
          spacing: { after: 7 },
          tabs: [{ pos: width, val: 'right', leader: 'dot' }],
        },
        schema.text(`${node.textContent.trim()}\t1`, [
          link.create({ href: `#${anchors[i]}` }),
        ]),
      );
    });
  }

  async listStyles(): Promise<DocStyle[]> {
    const all = (await this.host.styles?.()) ?? [];
    const used = new Map<string, number>();
    for (const { node } of this.textblocks()) {
      const heading = node.attrs['heading'] as number | null;
      const id = heading
        ? `Heading${heading}`
        : ((node.attrs['styleId'] as string | null) ?? null);
      if (id) used.set(id, (used.get(id) ?? 0) + 1);
    }
    return all
      .filter(
        (st) => st.type === 'paragraph' && (!st.hidden || used.has(st.id)),
      )
      .map((st) => ({
        id: st.id,
        name: st.name,
        custom: st.custom,
        ...(st.basedOn ? { basedOn: st.basedOn } : {}),
        used: used.get(st.id) ?? 0,
      }));
  }

  /** A document paragraph style by id, or by name (any case), with its look. */
  private async resolveStyle(nameOrId: string): Promise<ResolvedStyle> {
    const all = ((await this.host.styles?.()) ?? []).filter(
      (st) => st.type === 'paragraph',
    );
    const wanted = nameOrId.trim().toLowerCase();
    const style =
      all.find((st) => st.id === nameOrId) ??
      all.find((st) => st.name.toLowerCase() === wanted) ??
      all.find((st) => st.id.toLowerCase() === wanted);
    const look = style ? await this.host.styleFormatting?.(style.id) : null;
    if (!style || !look) {
      const names = all
        .filter((st) => !st.hidden)
        .map((st) => st.name)
        .slice(0, 20);
      throw new ContentError(
        `No paragraph style ${JSON.stringify(nameOrId)} in this document.` +
          (names.length
            ? ` It has: ${names.map((n) => JSON.stringify(n)).join(', ')}${all.length > names.length ? ', …' : ''} (list_styles).`
            : ' Title and Subtitle are always there.'),
      );
    }
    return { id: style.id, attrs: look.attrs, marks: look.marks };
  }

  /** Resolve every document style `content` names (by the name used). */
  private async stylesNamedIn(
    content: Content,
  ): Promise<Record<string, ResolvedStyle> | undefined> {
    if (typeof content === 'string') return undefined;
    const names = new Set<string>();
    for (const b of content)
      if (typeof b === 'object' && 'paragraph' in b && b.style)
        if (!isBuiltInStyle(b.style)) names.add(b.style);
    if (names.size === 0) return undefined;
    const out: Record<string, ResolvedStyle> = {};
    for (const name of names) out[name] = await this.resolveStyle(name);
    return out;
  }

  async save(): Promise<void> {
    await this.host.save();
  }

  async close(): Promise<void> {
    /* sessions over a host hold no resources of their own */
  }

  async editTable(
    tableIndex: number,
    edit: TableEdit,
    opts: MutationOptions = {},
  ): Promise<MutationResult & { rows: number; cols: number }> {
    this.checkVersion(opts.expectedVersion);
    const state = this.host.getState();
    const { schema } = state;
    const width = this.contentWidth();
    let tr = state.tr;
    const locate = () => {
      let found: { node: PMNode; pos: number } | null = null;
      let n = 0;
      tr.doc.descendants((node, pos) => {
        if (found) return false;
        if (node.type.name === 'table') {
          if (n === tableIndex) found = { node, pos };
          n++;
        }
        return !found;
      });
      if (!found) {
        throw new ContentError(
          `No table ${tableIndex} — get_document marks the blocks inside tables with table.index (0-based); the document has ${n}.`,
        );
      }
      return found as { node: PMNode; pos: number };
    };
    const rowAt = (table: PMNode, pos: number, r: number) => {
      if (r < 0 || r >= table.childCount) {
        throw new ContentError(
          `Row ${r} is out of range — the table has ${table.childCount} row(s).`,
        );
      }
      let rowPos = pos + 1;
      for (let i = 0; i < r; i++) rowPos += table.child(i).nodeSize;
      return { node: table.child(r), pos: rowPos };
    };

    if (edit.deleteTable) {
      const others = Object.entries(edit).filter(
        ([k, v]) => k !== 'deleteTable' && v !== undefined,
      );
      if (others.length > 0) {
        throw new ContentError(
          'delete_table takes the whole table away — pass it on its own.',
        );
      }
      const { node, pos } = locate();
      const $pos = state.doc.resolve(pos);
      // Whatever held the table (the page, a cell) must still hold a block.
      tr =
        $pos.parent.childCount === 1
          ? tr.replaceWith(
              pos,
              pos + node.nodeSize,
              schema.nodes['paragraph'].create(),
            )
          : tr.delete(pos, pos + node.nodeSize);
      this.host.apply(keepSections(state.doc, tr, 1));
      return {
        docVersion: this.host.getVersion(),
        range: { from: pos, to: pos },
        rows: 0,
        cols: 0,
      };
    }
    if (edit.deleteRows?.length) {
      const { node, pos } = locate();
      const rows = [...new Set(edit.deleteRows)].sort((a, b) => b - a);
      if (rows.length >= node.childCount)
        throw new ContentError(
          'A table must keep at least one row — pass delete_table to remove the whole table.',
        );
      for (const r of rows) {
        const row = rowAt(node, pos, r);
        tr = tr.delete(row.pos, row.pos + row.node.nodeSize);
      }
    }
    if (edit.deleteColumns?.length || edit.insertColumns) {
      tr = this.editColumns(tr, locate, edit, width, !edit.widths);
    }
    if (edit.insertRows) {
      const { node, pos } = locate();
      const grid = tableGrid(node, width);
      const at = edit.insertRows.at ?? node.childCount;
      if (at < 0 || at > node.childCount) {
        throw new ContentError(
          `Cannot insert at row ${at} — the table has ${node.childCount} row(s); omit at to append.`,
        );
      }
      const rows = edit.insertRows.rows.map((cells) =>
        rowNode(cells, grid.widths, schema),
      );
      const insertAt =
        at === node.childCount
          ? pos + node.nodeSize - 1
          : rowAt(node, pos, at).pos;
      tr = tr.insert(insertAt, rows);
    }
    if (edit.merge) {
      const { node, pos } = locate();
      const { row: r, from, to } = edit.merge;
      const row = rowAt(node, pos, r);
      if (from < 0 || to >= row.node.childCount || from > to) {
        throw new ContentError(
          `merge cells ${from}..${to} is out of range — row ${r} has ${row.node.childCount} cell(s).`,
        );
      }
      if (from < to) {
        let cellPos = row.pos + 1;
        for (let i = 0; i < from; i++) cellPos += row.node.child(i).nodeSize;
        let end = cellPos;
        let colspan = 0;
        const colwidth: number[] = [];
        let content = row.node.child(from).content;
        for (let i = from; i <= to; i++) {
          const cell = row.node.child(i);
          end += cell.nodeSize;
          colspan += Math.max(1, Number(cell.attrs['colspan']) || 1);
          const cw = cell.attrs['colwidth'] as number[] | null;
          if (cw) colwidth.push(...cw);
          if (i > from) content = content.append(cell.content);
        }
        const first = row.node.child(from);
        const merged = first.type.create(
          {
            ...first.attrs,
            colspan,
            colwidth: colwidth.length === colspan ? colwidth : null,
          },
          content,
        );
        tr = tr.replaceWith(cellPos, end, merged);
      }
    }
    if (edit.widths) {
      const { node, pos } = locate();
      const grid = tableGrid(node, width);
      tr = columnWidthsTr(
        tr,
        node,
        pos,
        columnWidths(edit.widths, grid.cols, width),
      );
    }
    if (edit.borders || edit.align !== undefined || edit.header !== undefined) {
      const { node, pos } = locate();
      const attrs: Record<string, unknown> = { ...node.attrs };
      const tableStyle = this.host.tableStyle?.();
      if (edit.borders) {
        const chrome = tableChrome(
          edit.borders,
          tableStyle,
          !!node.type.spec.attrs?.['styleId'],
        );
        attrs['styleId'] = chrome.styleId;
        attrs['look'] = chrome.look;
        attrs['borders'] = chrome.borders;
        if (
          chrome.styleId &&
          tableStyle?.style &&
          schema.nodes['doc'].spec.attrs?.['tableStyles']
        ) {
          const sheet = (tr.doc.attrs['tableStyles'] ?? {}) as Record<
            string,
            unknown
          >;
          if (!sheet[chrome.styleId]) {
            tr = tr.setDocAttribute('tableStyles', {
              ...sheet,
              [chrome.styleId]: tableStyle.style,
            });
          }
        }
      }
      if (edit.align !== undefined)
        attrs['align'] = edit.align === 'left' ? null : edit.align;
      tr = tr.setNodeMarkup(pos, undefined, attrs);
      if (edit.header !== undefined) {
        const first = rowAt(locate().node, pos, 0);
        tr = tr.setNodeMarkup(first.pos, undefined, {
          ...first.node.attrs,
          header: edit.header,
        });
      }
    }
    if (tr.steps.length === 0) {
      const { node } = locate();
      return {
        docVersion: this.host.getVersion(),
        rows: node.childCount,
        cols: tableGrid(node, width).cols,
      };
    }
    this.host.apply(tr);
    const { node, pos } = locate();
    return {
      docVersion: this.host.getVersion(),
      range: { from: pos, to: pos + node.nodeSize },
      rows: node.childCount,
      cols: tableGrid(node, width).cols,
    };
  }

  // ── internals ────────────────────────────────────────────────────────

  /**
   * Delete, then insert, grid columns of the table `locate` finds in `tr`.
   * Cells are addressed by the grid, merged cells included: a cell that spans
   * a deleted column narrows (it goes only when every column it spans does),
   * one that spans the insertion point widens, and a cell merged down from a
   * row above is changed once, in its own row. With `keepWidth` the columns
   * are rescaled afterwards so the table is as wide as it was.
   */
  private editColumns(
    tr: Transaction,
    locate: () => { node: PMNode; pos: number },
    edit: TableEdit,
    contentWidth: number,
    keepWidth: boolean,
  ): Transaction {
    const { schema } = tr.doc.type;
    const start = locate();
    let widths = tableGrid(start.node, contentWidth).widths;
    const total = widths.reduce((a, b) => a + b, 0);

    if (edit.deleteColumns?.length) {
      const { node, pos } = locate();
      const map = gridMap(node, pos);
      const doomed = new Set(edit.deleteColumns);
      for (const c of doomed) {
        if (c < 0 || c >= map.cols)
          throw new ContentError(
            `Column ${c} is out of range — the table has ${map.cols} column(s).`,
          );
      }
      if (doomed.size >= map.cols) {
        throw new ContentError(
          'A table must keep at least one column — pass delete_table to remove the whole table.',
        );
      }
      const ops: { pos: number; run: (t: Transaction) => Transaction }[] = [];
      map.rows.forEach((row, r) => {
        let kept = 0;
        for (const cell of row.cells) {
          let hit = 0;
          for (let c = cell.start; c < cell.start + cell.span; c++)
            if (doomed.has(c)) hit++;
          if (hit === 0) {
            kept++;
            continue;
          }
          if (hit === cell.span) {
            ops.push({
              pos: cell.pos,
              run: (t) => t.delete(cell.pos, cell.pos + cell.node.nodeSize),
            });
            continue;
          }
          kept++;
          const cw = cell.node.attrs['colwidth'] as number[] | null;
          ops.push({
            pos: cell.pos,
            run: (t) =>
              t.setNodeMarkup(cell.pos, undefined, {
                ...cell.node.attrs,
                colspan: cell.span - hit,
                colwidth:
                  cw && cw.length === cell.span
                    ? cw.filter((_w, i) => !doomed.has(cell.start + i))
                    : null,
              }),
          });
        }
        if (row.cells.length > 0 && kept === 0) {
          throw new ContentError(
            `Deleting those columns would leave row ${r} with no cell of its own (the rest of it is merged from ` +
              'the row above). Delete that row too, or fewer columns.',
          );
        }
      });
      for (const op of ops.sort((a, b) => b.pos - a.pos)) tr = op.run(tr);
      widths = widths.filter((_w, i) => !doomed.has(i));
    }

    if (edit.insertColumns) {
      const { node, pos } = locate();
      const map = gridMap(node, pos);
      const n = edit.insertColumns.count ?? 1;
      const at = edit.insertColumns.at ?? map.cols;
      if (!Number.isInteger(at) || at < 0 || at > map.cols) {
        throw new ContentError(
          `Cannot insert columns at ${at} — the table has ${map.cols} column(s); omit at to append.`,
        );
      }
      const w = Math.max(
        1,
        Math.round(widths.reduce((a, b) => a + b, 0) / widths.length),
      );
      const blank = () =>
        schema.nodes['table_cell'].create(
          { colspan: 1, colwidth: [w] },
          schema.nodes['paragraph'].create(),
        );
      const ops: { pos: number; run: (t: Transaction) => Transaction }[] = [];
      map.rows.forEach((row, r) => {
        const left = at > 0 ? map.occ[r][at - 1] : undefined;
        if (left && at < map.cols && left === map.occ[r][at]) {
          // A merged cell spans the insertion point: it widens, in its own row.
          if (left.row !== r) return;
          const cw = left.node.attrs['colwidth'] as number[] | null;
          const k = at - left.start;
          ops.push({
            pos: left.pos,
            run: (t) =>
              t.setNodeMarkup(left.pos, undefined, {
                ...left.node.attrs,
                colspan: left.span + n,
                colwidth:
                  cw && cw.length === left.span
                    ? [...cw.slice(0, k), ...Array(n).fill(w), ...cw.slice(k)]
                    : null,
              }),
          });
          return;
        }
        const next = row.cells.find((c) => c.start >= at);
        const insertAt = next ? next.pos : row.pos + row.node.nodeSize - 1;
        ops.push({
          pos: insertAt,
          run: (t) =>
            t.insert(
              insertAt,
              Array.from({ length: n }, () => blank()),
            ),
        });
      });
      for (const op of ops.sort((a, b) => b.pos - a.pos)) tr = op.run(tr);
      widths = [
        ...widths.slice(0, at),
        ...Array(n).fill(w),
        ...widths.slice(at),
      ];
    }

    if (keepWidth) {
      const sum = widths.reduce((a, b) => a + b, 0);
      const scale = sum > 0 ? total / sum : 1;
      const { node, pos } = locate();
      tr = columnWidthsTr(
        tr,
        node,
        pos,
        widths.map((x) => Math.max(1, Math.round(x * scale))),
      );
    }
    return tr;
  }

  /** 1-based section of a top-level block index, when the document has
   *  more than one section; null otherwise (nothing to tell apart). */
  private sectionOfTopLevel(): ((topIndex: number) => number) | null {
    const sections = this.host.getState().doc.attrs['sections'] as
      | { blockCount: number }[]
      | null
      | undefined;
    if (!sections || sections.length < 2) return null;
    const ends: number[] = [];
    let end = 0;
    for (const s of sections) ends.push((end += s.blockCount));
    return (i) => {
      const k = ends.findIndex((e) => i < e);
      return (k < 0 ? sections.length - 1 : k) + 1;
    };
  }

  /** The document's numbering definitions (`doc.attrs.numbering`). */
  private numbering(): NumberingDefs | null {
    return (
      (this.host.getState().doc.attrs['numbering'] as NumberingDefs | null) ??
      null
    );
  }

  /** The list a paragraph right before an insertion point belongs to — the
   *  list new items of the same kind join (so numbers count on). */
  private listBefore(
    node: PMNode | null | undefined,
  ): { kind: ListKind; numId: string } | null {
    const list = node ? listOf(node) : null;
    if (!list) return null;
    return {
      kind: listKindOf(list.numId, list.level, this.numbering()),
      numId: list.numId,
    };
  }

  /** The paragraph attrs (and numbering, when a list is born) for a
   *  Formatting's `list` / `listLevel` on the paragraph at `pos`. The indent
   *  moves with the level the way Tab moves it in the editor. */
  private listChange(
    node: PMNode,
    pos: number,
    format: Formatting,
  ): { attrs: Record<string, unknown>; numbering: NumberingDefs | null } {
    const current = listOf(node);
    const defs = this.numbering();
    const indent =
      (node.attrs['indent'] as Record<string, number> | null) ?? null;
    const shifted = (from: number, to: number): Record<string, unknown> => {
      if (from === to) return {};
      const left = Math.max(
        0,
        (indent?.['left'] ?? 0) + (to - from) * LIST_LEVEL_INDENT,
      );
      const next: Record<string, number> = { ...(indent ?? {}), left };
      if (left === 0) delete next['left'];
      return { indent: Object.keys(next).length > 0 ? next : null };
    };

    if (format.list === null) {
      if (!current) return { attrs: {}, numbering: null };
      return {
        attrs: { list: null, ...shifted(current.level, 0) },
        numbering: null,
      };
    }
    const kind =
      format.list ??
      (current ? listKindOf(current.numId, current.level, defs) : undefined);
    if (!kind) {
      throw new ContentError(
        'That paragraph is not a list item — pass list ("bullet" or "number") to make it one.',
      );
    }
    let numId: string;
    let numbering: NumberingDefs | null = null;
    if (current && listKindOf(current.numId, current.level, defs) === kind) {
      numId = current.numId;
    } else {
      const above = this.listBefore(
        this.host.getState().doc.resolve(pos).nodeBefore,
      );
      if (above?.kind === kind) numId = above.numId;
      else {
        numbering = { ...(defs ?? {}) };
        numId = mintList(kind, numbering);
      }
    }
    const from = current?.level ?? 0;
    const to =
      format.listLevel !== undefined
        ? this.listLevelFor(numId, format.listLevel, numbering ?? defs)
        : from;
    return {
      attrs: { list: { numId, level: to }, ...shifted(from, to) },
      numbering,
    };
  }

  /** A 1-based level an agent asked for → 0-based, within what the list's
   *  definition defines (the editor's lists: three). */
  private listLevelFor(
    numId: string,
    level: number,
    defs: NumberingDefs | null,
  ): number {
    const levels = defs?.[numId]?.levels;
    const defined = levels ? Object.keys(levels).map(Number) : [];
    const max = defined.length > 0 ? Math.max(...defined) + 1 : 9;
    if (!Number.isInteger(level) || level < 1 || level > max) {
      throw new ContentError(
        `List level ${level} is out of range — this list goes 1 to ${max} levels deep.`,
      );
    }
    return level - 1;
  }

  /** Resolve a formatting target to absolute positions. */
  private formatHit(
    target: Exclude<FormatTarget, { fromBlock: number }>,
    occurrence?: number,
  ): Hit {
    if (typeof target === 'string') return this.uniqueHit(target, occurrence);
    const blocks = this.textblocks();
    const block = blocks[target.blockIndex];
    if (!block) {
      throw new AnchorError(
        `blockIndex ${target.blockIndex} is out of range — the document has ${blocks.length} block(s).`,
      );
    }
    return {
      from: block.pos + 1,
      to: block.pos + 1 + block.node.content.size,
      blockIndex: target.blockIndex,
      context: block.node.textContent,
    };
  }

  /** The text area's width in px, from the document's page setup. */
  private contentWidth(): number {
    const page = this.host.getState().doc.attrs['page'] as
      | {
          width: number;
          margin: { left: number; right: number };
          gutter?: number;
        }
      | null
      | undefined;
    if (!page) return DEFAULT_CONTENT_WIDTH;
    return (
      page.width - page.margin.left - page.margin.right - (page.gutter ?? 0)
    );
  }

  private checkVersion(expected?: string): void {
    const current = this.host.getVersion();
    if (expected !== undefined && expected !== current) {
      throw new VersionConflictError(current, expected);
    }
  }

  /** All textblocks (paragraphs, incl. inside table cells) in reading order,
   *  each knowing which table cell holds it. */
  private textblocks(): TextBlock[] {
    return textblocksOf(this.host.getState().doc);
  }

  /** Every occurrence of `query`, atom-safe (matches never span images/fields
   *  or block boundaries), in document order with absolute PM positions. */
  private hits(query: string): Hit[] {
    return hitsIn(this.textblocks(), query);
  }

  /** Where an {@link InsertAnchor} puts new blocks, in PM positions. */
  private anchorPos(anchor: InsertAnchor): number {
    if (anchor.position === 'document_end')
      return this.host.getState().doc.content.size;
    const hit = this.uniqueHit(anchor.text, anchor.occurrence);
    const block = this.textblocks()[hit.blockIndex];
    return anchor.position === 'before'
      ? block.pos
      : block.pos + block.node.nodeSize;
  }

  /** The image a (blockIndex, imageIndex) pair from the latest snapshot
   *  addresses, with its absolute position. */
  private imageAt(
    blockIndex: number,
    imageIndex: number,
  ): { node: PMNode; pos: number } {
    const blocks = this.textblocks();
    const block = blocks[blockIndex];
    if (!block) {
      throw new AnchorError(
        `blockIndex ${blockIndex} is out of range — the document has ${blocks.length} block(s). ` +
          `Block indexes change with every edit; call get_document again.`,
      );
    }
    const images = blockImages(block.node).map((img) => ({
      node: img.node,
      pos: block.pos + 1 + img.offset,
    }));
    if (images.length === 0) {
      throw new AnchorError(
        `Block ${blockIndex} has no images. get_document lists each block's images.`,
      );
    }
    const img = images[imageIndex];
    if (!img) {
      throw new AnchorError(
        `imageIndex ${imageIndex} is out of range — block ${blockIndex} has ${images.length} image(s) (0-${images.length - 1}).`,
      );
    }
    return img;
  }

  /** The host's bytes for a source, with what the bytes say they are. */
  private async fetchImage(source: ImageSource): Promise<{
    bytes: Uint8Array;
    mediaType: string;
    width: number;
    height: number;
  }> {
    const offered = this.capabilities.images ?? [];
    if (!this.host.readImage || !offered.includes(source.kind)) {
      throw new ContentError(
        `This app cannot take a picture from ${source.kind}. ` +
          (offered.length
            ? `It accepts: ${offered.join(', ')}.`
            : 'It accepts no new pictures at all — tell the user to add the picture themselves.'),
      );
    }
    const got = await this.host.readImage(source);
    const info = sniffImage(got.bytes, got.mediaType);
    const scale = got.scale && got.scale > 0 ? got.scale : 1;
    return {
      bytes: got.bytes,
      mediaType: info.mediaType,
      width: info.width / scale,
      height: info.height / scale,
    };
  }

  /** The box a picture displays at: the asked-for width (or its own, shrunk
   *  to the text width when it is wider), with the height kept in proportion
   *  so nothing is ever squashed. */
  private boxFor(
    picture: { width: number; height: number },
    width?: number,
  ): ImageBox {
    const natural = Math.max(1, picture.width);
    const naturalHeight = Math.max(1, picture.height);
    const w = Math.max(
      1,
      Math.round(width ?? Math.min(natural, this.contentWidth())),
    );
    return {
      width: w,
      height: Math.max(1, Math.round((w * naturalHeight) / natural)),
    };
  }

  private uniqueHit(text: string, occurrence?: number): Hit {
    return pickHit(this.hits(text), text, occurrence, 'the document');
  }
}

/** All textblocks of `doc` (paragraphs, incl. inside table cells) in reading
 *  order, each knowing which table cell holds it. */
/**
 * The blocks a selection `from`–`to` covers, numbered the way get_document
 * numbers them (0-based, table-cell paragraphs included, reading order) —
 * the address an agent's tools take. A host names a user's selection with
 * this rather than by its own counting: top-level paragraph numbers drift
 * from the tools' numbering at the first table. `partial`: the selection
 * starts or ends inside a block's text, so block-wide formatting would
 * reach past it.
 *
 * Always a range: a selection that holds no text (a picture between
 * blocks, a caret) is named by the block it sits at — the one starting
 * at or before `from`, else the first — so a host never has a pick it
 * cannot point the agent to.
 */
export function selectionBlocks(
  doc: PMNode,
  from: number,
  to: number,
): { fromBlock: number; toBlock: number; partial: boolean } {
  let first = -1;
  let last = -1;
  let startsInside = false;
  let endsInside = false;
  let at = 0;
  textblocksOf(doc).forEach((b, i) => {
    if (b.pos <= from) at = i;
    const start = b.pos + 1;
    const end = b.pos + b.node.nodeSize - 1;
    const overlaps =
      Math.max(start, from) < Math.min(end, to) ||
      (start === end && from <= start && end <= to);
    if (!overlaps) return;
    if (first < 0) {
      first = i;
      startsInside = from > start;
    }
    last = i;
    endsInside = to < end;
  });
  return first < 0
    ? { fromBlock: at, toBlock: at, partial: true }
    : { fromBlock: first, toBlock: last, partial: startsInside || endsInside };
}

function textblocksOf(doc: PMNode): TextBlock[] {
  const out: TextBlock[] = [];
  let tables = 0;
  const walk = (node: PMNode, base: number, ctx: TextBlock['table']) => {
    node.forEach((child, offset, i) => {
      const pos = base + offset;
      if (child.isTextblock) {
        out.push(ctx ? { node: child, pos, table: ctx } : { node: child, pos });
        return;
      }
      let next = ctx;
      if (child.type.name === 'table')
        next = { index: tables++, row: 0, cell: 0 };
      else if (child.type.name === 'table_row' && ctx)
        next = { ...ctx, row: i, cell: 0 };
      else if (child.type.name === 'table_cell' && ctx)
        next = { ...ctx, cell: i };
      walk(child, pos + 1, next);
    });
  };
  walk(doc, 0, undefined);
  return out;
}

/** Every occurrence of `query` in `blocks`, atom-safe: matches never span
 *  images, fields or block boundaries. */
function hitsIn(blocks: TextBlock[], query: string): Hit[] {
  if (query.length === 0) return [];
  const out: Hit[] = [];
  blocks.forEach(({ node, pos }, blockIndex) => {
    // Concatenate the block's text children, breaking the searchable string
    // at non-text inlines so a match can't pretend to span an atom. Each
    // segment records where it starts in the joined string AND in the doc.
    const segments: {
      joinedStart: number;
      length: number;
      startPos: number;
    }[] = [];
    let joined = '';
    node.forEach((child, offset) => {
      if (child.isText && child.text) {
        segments.push({
          joinedStart: joined.length,
          length: child.text.length,
          startPos: pos + 1 + offset,
        });
        joined += child.text;
      } else {
        joined += '￿'; // unmatchable atom sentinel
      }
    });
    let at = joined.indexOf(query);
    while (at !== -1) {
      // A match may span adjacent text segments (mark changes split runs);
      // adjacency in `joined` implies adjacency in the doc, so mapping the
      // START offset to a position is enough.
      const seg = segments.find(
        (s) => at >= s.joinedStart && at < s.joinedStart + s.length,
      );
      if (seg) {
        const from = seg.startPos + (at - seg.joinedStart);
        out.push({
          from,
          to: from + query.length,
          blockIndex,
          context: contextAround(joined.replace(/￿/g, ' '), at, query.length),
        });
      }
      at = joined.indexOf(query, at + 1);
    }
  });
  return out;
}

/** The one hit an anchor means: unique, or picked by occurrence — with the
 *  sentence that teaches the retry when it is neither. */
function pickHit(
  all: Hit[],
  text: string,
  occurrence: number | undefined,
  where: string,
): Hit {
  if (all.length === 0) {
    throw new AnchorError(
      `Text not found in ${where}: ${JSON.stringify(clip(text))}. ` +
        `Anchors match within one paragraph — check get_document for the exact text.`,
    );
  }
  if (occurrence !== undefined) {
    const hit = all[occurrence - 1];
    if (!hit) {
      throw new AnchorError(
        `occurrence ${occurrence} is out of range — the text matches ${all.length} time(s).`,
      );
    }
    return hit;
  }
  if (all.length > 1) {
    throw new AnchorError(
      `The text matches ${all.length} times — pass occurrence (1-${all.length}) to pick one, ` +
        `or use a longer, unique anchor.`,
    );
  }
  return all[0];
}

/** What an image box actually holds — an agent about to swap one out needs
 *  to know whether it is a photo or art the file describes shape by shape. */
function imageKind(img: PMNode): DocImageKind {
  const a = img.attrs;
  if (a['oleProgId'] || a['equation'] || a['equationAst']) return 'equation';
  if (a['rawDrawing']) return 'drawing';
  if (a['shape'] || a['vector']) return 'shape';
  return 'bitmap';
}

/** The block's inline image children, in order, with their child offsets. */
function blockImages(block: PMNode): { node: PMNode; offset: number }[] {
  const out: { node: PMNode; offset: number }[] = [];
  block.forEach((child, offset) => {
    if (child.type.name === 'image') out.push({ node: child, offset });
  });
  return out;
}

/** Each table of contents in the body: its run of consecutive top-level
 *  paragraphs sharing one TOC field object (the model's fieldAt rule). */
function tocSpans(doc: PMNode): {
  field: { kind: string; instr: string };
  from: number;
  to: number;
  count: number;
}[] {
  const spans: {
    field: { kind: string; instr: string };
    from: number;
    to: number;
    count: number;
  }[] = [];
  doc.forEach((node, offset) => {
    const f = node.attrs['field'] as { kind?: string; instr?: string } | null;
    const last = spans[spans.length - 1];
    if (f && last && last.field === f && last.to === offset) {
      last.to = offset + node.nodeSize;
      last.count++;
    } else if (f?.kind === 'toc' && f.instr) {
      spans.push({
        field: f as { kind: string; instr: string },
        from: offset,
        to: offset + node.nodeSize,
        count: 1,
      });
    }
  });
  return spans;
}

/** The heading levels a TOC instruction lists — `\\o "1-3"` → 3. */
function tocLevels(instr: string): number {
  const m = /\\o\s+"\d+-(\d+)"/.exec(instr);
  const n = m ? Number(m[1]) : 3;
  return n >= 1 && n <= 9 ? n : 3;
}

/**
 * Keep `doc.attrs.sections` where it was across a transaction that adds or
 * removes top-level blocks.
 *
 * Sections are stored by COUNT of top-level blocks, so a block added above a
 * section break, or taken away, moves every break below it — a landscape
 * section would start a paragraph early. The boundaries (the positions
 * between the last block of one section and the first of the next) are
 * mapped through the transaction instead, and the counts re-read from where
 * they land. `assoc` decides a block added exactly at a boundary: 1 keeps it
 * in the section above (inserted after that section's last block), -1 in the
 * one below (inserted before its first block).
 *
 * A change that would empty a section is refused: the break would go with it,
 * and whatever is keyed by section (headers, page numbers) would shift.
 */
export function keepSections(
  before: PMNode,
  tr: Transaction,
  assoc: -1 | 1,
): Transaction {
  const sections = before.attrs['sections'] as
    | ({ blockCount: number } & Record<string, unknown>)[]
    | null
    | undefined;
  if (!sections || sections.length < 2 || !tr.docChanged) return tr;
  const offsets: number[] = []; // start offset of every top-level block, + the end
  before.forEach((_child, offset) => offsets.push(offset));
  offsets.push(before.content.size);
  const after: number[] = [];
  tr.doc.forEach((_child, offset) => after.push(offset));
  after.push(tr.doc.content.size);

  let start = 0;
  let prevIndex = 0;
  let changed = false;
  const next = sections.map((s, i) => {
    start += s.blockCount;
    const last = i === sections.length - 1;
    const index = last
      ? tr.doc.childCount
      : after.findIndex(
          (o) =>
            o >=
            tr.mapping.map(offsets[Math.min(start, offsets.length - 1)], assoc),
        );
    const blockCount = (index < 0 ? tr.doc.childCount : index) - prevIndex;
    prevIndex += blockCount;
    if (blockCount <= 0) {
      throw new ContentError(
        `That would empty section ${i + 1} of the document, taking its section break with it. ` +
          'Leave at least one of its paragraphs, or ask the user to remove the section break.',
      );
    }
    if (blockCount !== s.blockCount) changed = true;
    return blockCount === s.blockCount ? s : { ...s, blockCount };
  });
  return changed ? tr.setDocAttribute('sections', next) : tr;
}

/** One cell on a table's grid: where it starts, how many columns it spans,
 *  and the row it belongs to (a cell merged down covers rows below it). */
interface GridCell {
  node: PMNode;
  pos: number;
  row: number;
  start: number;
  span: number;
}

/** A table laid out on its grid: every row's own cells, and for every
 *  (row, column) the cell that covers it — including one merged down from
 *  a row above, which that row does not contain. */
function gridMap(
  table: PMNode,
  tablePos: number,
): {
  cols: number;
  rows: { node: PMNode; pos: number; cells: GridCell[] }[];
  occ: (GridCell | undefined)[][];
} {
  const rows: { node: PMNode; pos: number; cells: GridCell[] }[] = [];
  const occ: (GridCell | undefined)[][] = [];
  let cols = 0;
  let rowPos = tablePos + 1;
  table.forEach((row, _offset, r) => {
    const here = (occ[r] ??= []);
    const cells: GridCell[] = [];
    let col = 0;
    let cellPos = rowPos + 1;
    row.forEach((cell) => {
      while (here[col]) col++;
      const span = Math.max(1, Number(cell.attrs['colspan']) || 1);
      const down = Math.max(1, Number(cell.attrs['rowspan']) || 1);
      const g: GridCell = {
        node: cell,
        pos: cellPos,
        row: r,
        start: col,
        span,
      };
      cells.push(g);
      for (let rr = r; rr < Math.min(r + down, table.childCount); rr++) {
        const line = (occ[rr] ??= []);
        for (let c = col; c < col + span; c++) line[c] = g;
      }
      col += span;
      cellPos += cell.nodeSize;
    });
    cols = Math.max(cols, here.length);
    rows.push({ node: row, pos: rowPos, cells });
    rowPos += row.nodeSize;
  });
  return { cols, rows, occ };
}

/** Give every cell of the table at `tablePos` the widths (px, one per grid
 *  column) of the columns it spans — by the grid, so a row that continues a
 *  cell merged from above lines up with the rest. */
function columnWidthsTr(
  tr: Transaction,
  table: PMNode,
  tablePos: number,
  widths: number[],
): Transaction {
  for (const row of gridMap(table, tablePos).rows) {
    for (const cell of row.cells) {
      tr = tr.setNodeMarkup(cell.pos, undefined, {
        ...cell.node.attrs,
        colwidth: widths.slice(cell.start, cell.start + cell.span),
      });
    }
  }
  return tr;
}

/** The hyperlinks of a block, adjacent runs to one address read as one. */
function blockLinks(block: PMNode): { text: string; href: string }[] {
  const out: { text: string; href: string }[] = [];
  let open: { text: string; href: string } | null = null;
  block.forEach((child) => {
    const href = child.marks.find((m) => m.type.name === 'link')?.attrs[
      'href'
    ] as string | undefined;
    if (!href || !child.isText) {
      open = null;
      return;
    }
    if (open && open.href === href) open.text += child.text ?? '';
    else {
      open = { text: child.text ?? '', href };
      out.push(open);
    }
  });
  return out;
}

/** The footnote references of `doc` in order: num, the number shown, pos. */
function footnoteRefs(
  doc: PMNode,
): { num: number; shown: string; pos: number }[] {
  const out: { num: number; shown: string; pos: number }[] = [];
  doc.descendants((n, pos) => {
    if (!n.isText) return true;
    const fn = n.marks.find((m) => m.type.name === 'footnote');
    if (fn)
      out.push({ num: Number(fn.attrs['num']), shown: n.text ?? '', pos });
    return false;
  });
  return out;
}

/** How a page field reads in chrome text: where the number goes. */
function fieldText(leaf: PMNode): string {
  if (leaf.type.name !== 'page_field') return '';
  return leaf.attrs['kind'] === 'pages' ? '{pages}' : '{page}';
}

/** `story` with the text at `hit` replaced, in the marks of its first
 *  character. Built from the story's own nodes (this tier carries no
 *  prosemirror-transform): the new text travels as a slice of a scratch
 *  paragraph, open on both sides, so it joins the paragraph it lands in. */
function replaceInStory(story: PMNode, hit: Hit, newText: string): PMNode {
  const schema = story.type.schema;
  if (!newText)
    return story.replace(hit.from, hit.to, story.slice(hit.from, hit.from));
  const marks = story.nodeAt(hit.from)?.marks ?? [];
  const scratch = schema.node('doc', null, [
    schema.node('paragraph', null, [schema.text(newText, marks)]),
  ]);
  return story.replace(hit.from, hit.to, scratch.slice(1, 1 + newText.length));
}

/** A paragraph's list membership (0-based level), or null for body text. */
function listOf(node: PMNode): { numId: string; level: number } | null {
  const list = node.attrs['list'] as
    | { numId?: string; level?: number }
    | null
    | undefined;
  if (!list?.numId) return null;
  return { numId: list.numId, level: list.level ?? 0 };
}

/** A heading, a Title or a Subtitle — what apply_formatting's
 *  only: "headings" picks, and "body" leaves out. */
function isHeadingBlock(node: PMNode): boolean {
  const heading = node.attrs['heading'] as number | null | undefined;
  if (typeof heading === 'number' && heading >= 1) return true;
  const style = node.attrs['styleId'] as string | null | undefined;
  return style === 'Title' || style === 'Subtitle';
}

function blockType(node: PMNode): string {
  const heading = node.attrs['heading'] as number | null | undefined;
  if (typeof heading === 'number' && heading >= 1) return `heading${heading}`;
  return node.type.name;
}

function contextAround(text: string, at: number, len: number): string {
  const before = text.slice(Math.max(0, at - 30), at);
  const after = text.slice(at + len, at + len + 30);
  return `${before}«${text.slice(at, at + len)}»${after}`;
}

function clip(s: string): string {
  return s.length > 60 ? `${s.slice(0, 57)}…` : s;
}
