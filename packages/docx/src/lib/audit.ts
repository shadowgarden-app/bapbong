import type { OoxmlNode } from './ooxml.js';

/**
 * Flag-gated XML **coverage audit**: when enabled, an import logs every tag and
 * attribute that existed in the .docx but was never read by the converter, and
 * an export logs every ProseMirror node/mark type the serializer didn't handle.
 * This is a measurement-only tool for finding conversion gaps — it changes no
 * behavior and costs one boolean check per accessor call when off.
 *
 * Enable it at runtime with either:
 *   - `globalThis.__BAPBONG_XML_AUDIT__ = true`
 *   - `localStorage.setItem('bapbong.xmlAudit', '1')`  (browser / webview)
 *
 * How it measures (import): every read of the parsed OOXML tree goes through
 * the ooxml.ts accessors (`child`/`children`/`attrOf`/`findDescendant`), which
 * mark the nodes/attributes they touch. A few walkers that iterate
 * `node.children` directly mark the branches they actually consume. After the
 * import, the whole tree of every registered part is walked and anything
 * untouched is reported — so the report reflects what the code *really does*,
 * with no hand-maintained support list to drift.
 *
 * Reporting rules:
 *   - An untouched element is reported once, only when its PARENT was touched
 *     (the boundary of coverage); deeper untouched descendants are implied.
 *     Touched descendants under an untouched ancestor still get attr-checked.
 *   - `markSubtree` covers a node AND everything under it — for branches
 *     consumed wholesale by design (dropped w:del content, flattened OMML,
 *     comment bodies reduced to plain text).
 *   - An attribute counts as covered when the code ASKED for it (`attrOf`),
 *     even if absent in this file — checking is coverage.
 *   - Entries are split three ways: UNKNOWN (real gaps), inert (unread, but
 *     at a value the spec defines as a no-op — see INERT_* below), and
 *     ignored-by-design (the IGNORED_* lists: revision noise, style-gallery
 *     metadata, …).
 *
 * Known limitation: a read attribute whose VALUE the code didn't recognize
 * (e.g. `w:jc w:val="thaiDistribute"`) is not caught — the attr was read.
 *
 * Overlapping imports (the playground parses a file twice concurrently) are
 * depth-counted into ONE merged report — marks live on per-parse node
 * objects, so the merge is safe; counts just sum across the overlapping
 * parses. importDocx pairs begin/end in try/finally, so a throw mid-import
 * can't wedge the depth.
 */

export interface AuditEntry {
  /** Part path, e.g. "word/document.xml". Export entries use "model". */
  part: string;
  /** "w:keepNext" for a tag, "w:pgMar @w:header" for an attribute,
   *  "mark comment" / "node foo" for export-side model types. */
  key: string;
  count: number;
}

/** An entry mid-collection, before bucketing: same as AuditEntry plus the
 *  value-derived verdict only the tree walk can compute. */
interface Counted extends AuditEntry {
  inert: boolean;
}

export interface AuditReport {
  mode: 'import' | 'export';
  label: string;
  unknown: AuditEntry[];
  /** Unread, but carrying a value the spec defines as having no effect — so
   *  reading it would change nothing on the page. See INERT_ATTRS. */
  inert: AuditEntry[];
  ignored: AuditEntry[];
}

/** Tags skipped on purpose — being unread is a decision, not a gap. */
const IGNORED_TAGS = new Set([
  // Revision/proofing noise Word sprinkles everywhere.
  'w:proofErr',
  'w:noProof',
  'w:lang',
  'w:lastRenderedPageBreak',
  // The paired end marker of a bookmark: the START carries name + id and IS
  // read (paragraph `bookmarks` attr); the end is a bare id we regenerate.
  'w:bookmarkEnd',
  // Style-gallery metadata: affects Word's styles UI, never rendering.
  'w:name',
  'w:next',
  'w:link',
  'w:qFormat',
  'w:semiHidden',
  'w:uiPriority',
  'w:unhideWhenUsed',
  'w:aliases',
  'w:autoRedefine',
  'w:locked',
  'w:rsid',
  // Complex-script twins of properties we do read (w:sz/w:b/w:i).
  'w:szCs',
  'w:bCs',
  'w:iCs',
  // Comment machinery we deliberately skip: the range markers carry the
  // anchor; the reference glyph run adds nothing.
  'w:commentReference',
  // Style-gallery / revision bookkeeping in styles.xml & settings.xml.
  'w:latentStyles',
  'w:rsids',
  'w:shapeDefaults',
  'w:themeFontLang',
  'w:listSeparator',
  'w:zoom',
  'w:clrSchemeMapping',
  'w:embedTrueTypeFonts',
  'w:proofState',
  // Drawing chrome around pictures (ids, stretch/crop boilerplate) — the
  // content (blip, extent, wrap, alt text) is read where it matters.
  'pic:nvPicPr',
  'pic:cNvPicPr',
  'a:stretch',
  'wp:effectExtent',
  'wp:cNvGraphicFramePr',
  // Non-visual shape/connector/group properties — pure editor metadata.
  'wps:cNvSpPr',
  'wps:cNvCnPr',
  'wpg:cNvGrpSpPr',
  // xfrm offsets/extents on single pictures mirror wp:extent (which IS
  // read); groups DO read their xfrm through parseGroup.
  'a:off',
  'a:ext',
  // Vendor extension lists and preset-geometry adjustment defaults.
  'a:extLst',
  'a:avLst',
  // Line-end/join/effect cosmetics on outlines we already paint solid.
  'a:headEnd',
  'a:tailEnd',
  'a:miter',
  'a:round',
  'a:lum',
  'a:noAutofit',
  // Shape style refs beyond a:lnRef (which IS read for outline color):
  // effects (shadows) and themed shape text fonts are out of paint scope.
  'a:effectRef',
  'a:fontRef',
  // Picture-frame geometry: Word writes prstGeom rect boilerplate on every
  // photo; wps SHAPE prstGeom is read and drives ShapeSpec.
  'a:prstGeom',
  // spPr background "no fill" boilerplate on pictures (a:ln noFill IS
  // absence-checked where outlines are painted).
  'a:noFill',
  // Anchor positioning we place by posOffset/align: simplePos is the
  // legacy fallback scheme, wrapNone is the fall-through of the wrap probe.
  'wp:simplePos',
  'wp:wrapNone',
  // Settings that tune Word's own editing/proofing UI, not rendering; the
  // whole settings.xml part is carried on export.
  'w:compat',
  'w:decimalSymbol',
  'w:characterSpacingControl',
  'w:hideSpellingErrors',
  'w:hideGrammaticalErrors',
  'w:displayBackgroundShape',
  'w:formProt',
  // Hyphenation is out of layout scope by decision.
  'w:autoHyphenation',
  'w:hyphenationZone',
  'w:suppressAutoHyphens',
  // OMML is flattened to plain text by design — math defaults are moot.
  'm:mathPr',
  // Tracked-change record of an older table grid — the current grid is read.
  'w:tblGridChange',
  // Per-cell width copies of the grid: w:tblGrid is the width authority
  // (Word keeps them in sync on save); pct-typed and grid-fallback widths
  // ARE read from the first row where they matter. Also re-emitted on
  // export from the model's colwidths, so nothing is lost.
  'w:tcW',
  // Footnote-body marker glyph: the layout draws its own note numbers.
  'w:footnoteRef',
  // Theme parts beyond colors, fonts and the format scheme: object defaults
  // and extra scheme variants style Office's galleries, not document content.
  // a:fmtScheme is NOT here any more — a shape's a:fillRef resolves through
  // it, so the entries we cannot paint (gradient fills, the theme's line
  // widths) are real gaps and should stay countable.
  'a:objectDefaults',
  'a:extraClrSchemeLst',
  // Theme effect gallery: we paint no shadows or glows at all, the same
  // decision already recorded for a:effectRef and a:effectLst.
  'a:effectStyleLst',
  // numbering.xml bookkeeping: internal ids and list-gallery metadata,
  // meaningless to rendering; the part itself is carried on export.
  'w:nsid',
  'w:multiLevelType',
  'w:tmpl',
  'w:numIdMacAtCleanup',
]);

/** Tag prefixes skipped on purpose (markup-compat wrappers, w14/w15 extras —
 *  the w15 comments-extended part we DO read is visited, so never reported). */
const IGNORED_TAG_PREFIXES = [
  'mc:',
  'w14:',
  'w15:',
  'w16cid:',
  'w16se:',
  'wp14:', // sizeRelH/V etc. — 2010 drawing extensions
];

function isIgnoredTag(name: string): boolean {
  if (IGNORED_TAGS.has(name)) return true;
  return IGNORED_TAG_PREFIXES.some((p) => name.startsWith(p));
}

/** Attribute names skipped on purpose (on otherwise-covered elements). */
function isIgnoredAttr(tag: string, name: string): boolean {
  return (
    name.startsWith('xmlns') ||
    name.startsWith('w:rsid') ||
    name.startsWith('w14:') ||
    name.startsWith('w15:') ||
    name.startsWith('wp14:') || // anchorId/editId — revision markers
    name.startsWith('w16cid:') || // durable ids — revision bookkeeping
    name === 'xml:space' ||
    name === 'mc:Ignorable' ||
    // List-gallery metadata on levels: tplc names the gallery template,
    // tentative flags levels Word invented but nothing uses yet.
    (tag === 'w:lvl' && (name === 'w:tplc' || name === 'w:tentative')) ||
    // Rels are resolved by Id; the relationship Type is package plumbing.
    (tag === 'Relationship' && name === 'Type') ||
    // Whether a style came from Word's built-in gallery or the user made it:
    // drives the Styles pane and built-in name mapping, never rendering.
    (tag === 'w:style' && name === 'w:customStyle') ||
    // Word's visited-link tracking flag — UI state, not content.
    (tag === 'w:hyperlink' && name === 'w:history') ||
    // A bookmark's numeric id pairs start/end within the part; the NAME is
    // what links point at, and the exporter renumbers on the way out.
    (tag === 'w:bookmarkStart' &&
      (name === 'w:id' || name === 'w:displacedByCustomXml')) ||
    // Drawing object ids/names are display metadata; export regenerates them.
    (tag === 'wp:docPr' && (name === 'id' || name === 'name')) ||
    // Same for VML: `id`/`o:spid` identify a shape within the part (OLE and
    // legacy form controls point at them); they name nothing and the exporter
    // writes DrawingML with fresh ids. The DESCRIPTION lives in `alt`, which
    // is read.
    (tag.startsWith('v:') && (name === 'id' || name === 'o:spid')) ||
    // Inline-drawing gaps: Word hard-codes 0 on wp:inline (the float-side
    // dist* on wp:anchor IS read and applied).
    (tag === 'wp:inline' && name.startsWith('dist')) ||
    // graphicData is dispatched by its CHILD tag (pic:pic / wps:wsp), not uri.
    (tag === 'a:graphicData' && name === 'uri') ||
    // Theme scheme names ("Office", …) — gallery labels, not content.
    (tag.startsWith('a:') && name === 'name') ||
    // AlternateContent is dispatched by CHILD content (pic/wps), not by the
    // Requires namespace list.
    (tag.startsWith('mc:') && name === 'Requires') ||
    // Drawing object ids/names and paint hints; export regenerates them.
    (tag === 'pic:cNvPr' && (name === 'id' || name === 'name')) ||
    (tag.endsWith(':spPr') && name === 'bwMode') ||
    (tag === 'a:blip' && name === 'cstate') ||
    (tag === 'pic:blipFill' && name === 'rotWithShape') ||
    (tag === 'a:lnRef' && name === 'idx') ||
    // Anchor toggles Word's own layouter consults; we place by offset/align.
    (tag === 'wp:anchor' &&
      [
        'allowOverlap',
        'layoutInCell',
        'locked',
        'relativeHeight',
        'simplePos',
      ].includes(name)) ||
    // CJK disambiguation hint — the explicit rFonts attrs are all read.
    (tag === 'w:rFonts' && name === 'w:hint')
  );
}

/**
 * "Unread" and "would change the page" are not the same question, and the gap
 * between them is almost entirely one thing: values that do nothing. Word
 * writes a full set of schema defaults on every textbox, an empty a:srcRect on
 * every picture, a "no outline" a:ln on every photo — none of which the
 * document asked for. Counting those as gaps buries the real ones.
 *
 * An entry lands in the `inert` bucket when it is unread AND, per ECMA-376,
 * a no-op in exactly this state. Two shapes:
 *   - an attribute sitting at its schema default (`wps:bodyPr @anchor="t"`);
 *   - an element the spec defines as inert here (`a:ln` holding only
 *     `a:noFill`, an empty `a:effectLst`, `prst="textNoShape"`).
 *
 * What does NOT qualify — the distinction this whole bucket rests on — is
 * "the value happens to equal our hardcoded fallback". Those render the same
 * only by coincidence and are real gaps: `w:tblCellMar`'s 108 twips matching
 * CELL_PAD_X's 7.2px is the standing example, and it stays UNKNOWN.
 *
 * Demoting by VALUE rather than by NAME is what keeps this from going blind:
 * any other value on the same key falls straight back to UNKNOWN, so the day
 * a file writes `anchor="ctr"` the audit says so.
 */
const FALSE_VALUES = new Set(['0', 'false', 'off']);
const isFalse = (v: string) => FALSE_VALUES.has(v);
const isNum = (n: number) => (v: string) => Number(v) === n;

/** `tag @attr` → predicate on the raw value: true = this value is a no-op. */
const INERT_ATTRS: Record<string, (v: string) => boolean> = {
  // CT_TextBodyProperties. Word stamps the whole default set onto every
  // textbox it writes. NOT listed: @compatLnSpc, whose default is false — a
  // written "1" really does change line spacing inside the shape.
  'wps:bodyPr @rot': isNum(0),
  'wps:bodyPr @spcFirstLastPara': isFalse,
  'wps:bodyPr @vertOverflow': (v) => v === 'overflow',
  'wps:bodyPr @horzOverflow': (v) => v === 'overflow',
  'wps:bodyPr @vert': (v) => v === 'horz',
  'wps:bodyPr @wrap': (v) => v === 'square',
  'wps:bodyPr @numCol': isNum(1),
  'wps:bodyPr @spcCol': isNum(0),
  'wps:bodyPr @rtlCol': isFalse,
  'wps:bodyPr @fromWordArt': isFalse,
  'wps:bodyPr @anchor': (v) => v === 't',
  'wps:bodyPr @anchorCtr': isFalse,
  'wps:bodyPr @forceAA': isFalse,
};

/** Elements that are no-ops in a particular shape. The predicate reads
 *  `node.attrs`/`node.children` directly — the sweep must not mark anything. */
const INERT_TAGS: Record<string, (n: OoxmlNode) => boolean> = {
  // CT_RelativeRect: all four insets default to 0, so an empty srcRect (what
  // Word writes on every uncropped picture) crops nothing.
  'a:srcRect': (n) => n.children.length === 0 && attrCount(n) === 0,
  // No effect children = no shadow, glow or reflection.
  'a:effectLst': (n) => n.children.length === 0,
  // The spec's explicit "this shape has no outline" state.
  'a:ln': (n) =>
    n.children.length > 0 && n.children.every((c) => c.name === 'a:noFill'),
  // The preset that applies no warp at all — Word writes it on every textbox.
  'a:prstTxWarp': (n) => n.attrs['prst'] === 'textNoShape',
  // ST_TextEffect defaults to "none": the animated text effects Word has not
  // rendered since 2007, written out as "no effect".
  'w:effect': (n) => n.attrs['w:val'] === 'none',
  // CT_TblWidth @w:w defaults to 0 = the table sits at the margin.
  'w:tblInd': (n) => Number(n.attrs['w:w'] ?? '0') === 0,
  // ST_DocGrid @w:type defaults to "default", which snaps nothing — and that
  // makes the linePitch/charSpace it carries moot. Only lines/linesAndChars
  // actually grid the page, and those stay UNKNOWN.
  'w:docGrid': (n) => {
    const t = n.attrs['w:type'];
    return t === undefined || t === 'default';
  },
};

function attrCount(n: OoxmlNode): number {
  return Object.keys(n.attrs).filter((a) => !a.startsWith('xmlns')).length;
}

function isInertAttr(tag: string, name: string, value: string): boolean {
  return INERT_ATTRS[`${tag} @${name}`]?.(value) ?? false;
}

function isInertTag(node: OoxmlNode): boolean {
  return INERT_TAGS[node.name]?.(node) ?? false;
}

function computeEnabled(): boolean {
  const g = globalThis as unknown as {
    __BAPBONG_XML_AUDIT__?: unknown;
    localStorage?: { getItem(k: string): string | null };
  };
  if (g.__BAPBONG_XML_AUDIT__ != null) return !!g.__BAPBONG_XML_AUDIT__;
  try {
    return g.localStorage?.getItem('bapbong.xmlAudit') === '1';
  } catch {
    return false;
  }
}

// ── session state ───────────────────────────────────────────────────
// `importActive` is the fast-path gate: begin() resolves the enabled flag ONCE
// per session so the hot accessors read one module boolean instead of hitting
// globalThis/localStorage per call. `importDepth` merges overlapping imports:
// only the outermost begin resets state and only the matching end emits.
let importActive = false;
let importDepth = 0;
let exportActive = false;
let importLabel = '';
let exportLabel = '';
let visited = new WeakSet<OoxmlNode>();
let subtrees = new WeakSet<OoxmlNode>();
let readAttrs = new WeakMap<OoxmlNode, Set<string>>();
let parts: { name: string; root: OoxmlNode }[] = [];
let exportCounts = new Map<string, number>();
let lastReport: AuditReport | null = null;

// Optional extra destination for report lines (same idea as perf.setSink) —
// for hosts whose console is invisible, e.g. the packaged desktop WKWebView.
let sink: ((line: string) => void) | null = null;

function emitLine(line: string): void {
  console.log(line);
  if (sink) {
    try {
      sink(line);
    } catch {
      // a broken sink must never break the audited code
    }
  }
}

function classify(counts: Map<string, Counted>): {
  unknown: AuditEntry[];
  inert: AuditEntry[];
  ignored: AuditEntry[];
} {
  const unknown: AuditEntry[] = [];
  // UNKNOWN entries keep their part (the location matters for fixing them);
  // inert and ignored-by-design ones aggregate across parts — pure noise
  // volume otherwise (the same rsid attr once per header/footer part).
  const inertByKey = new Map<string, AuditEntry>();
  const ignoredByKey = new Map<string, AuditEntry>();
  const aggregate = (into: Map<string, AuditEntry>, e: Counted) => {
    const agg = into.get(e.key);
    if (agg) agg.count += e.count;
    else into.set(e.key, { part: '*', key: e.key, count: e.count });
  };
  for (const e of counts.values()) {
    const tag = e.key.split(' @')[0];
    const attr = e.key.includes(' @') ? e.key.split(' @')[1] : null;
    if (attr ? isIgnoredAttr(tag, attr) : isIgnoredTag(tag)) {
      aggregate(ignoredByKey, e);
    } else if (e.inert) {
      aggregate(inertByKey, e);
    } else {
      unknown.push({ part: e.part, key: e.key, count: e.count });
    }
  }
  const byCount = (a: AuditEntry, b: AuditEntry) =>
    b.count - a.count || a.key.localeCompare(b.key);
  unknown.sort(byCount);
  const inert = [...inertByKey.values()].sort(byCount);
  const ignored = [...ignoredByKey.values()].sort(byCount);
  return { unknown, inert, ignored };
}

function emitReport(report: AuditReport): void {
  lastReport = report;
  const { mode, label, unknown, inert, ignored } = report;
  const header =
    `[xml-audit] ${mode} "${label}" — ` +
    `${unknown.length} UNKNOWN, ${inert.length} inert, ` +
    `${ignored.length} ignored-by-design`;
  const c = console as unknown as {
    groupCollapsed?: (l: string) => void;
    groupEnd?: () => void;
  };
  const grouped = typeof c.groupCollapsed === 'function';
  if (grouped) c.groupCollapsed?.(header);
  else emitLine(header);
  if (sink && grouped) sink(header);
  const pad = Math.max(
    0,
    ...[...unknown, ...inert, ...ignored].map((e) => e.part.length),
  );
  for (const e of unknown)
    emitLine(`  UNKNOWN ${e.part.padEnd(pad)}  ${e.key}  ×${e.count}`);
  for (const e of inert)
    emitLine(`  inert   ${e.part.padEnd(pad)}  ${e.key}  ×${e.count}`);
  for (const e of ignored)
    emitLine(`  ignored ${e.part.padEnd(pad)}  ${e.key}  ×${e.count}`);
  if (grouped) c.groupEnd?.();
}

/** Walk a registered part, collecting untouched tags/attrs (see module doc
 *  for the boundary rule). Inert-ness is decided HERE, the only place that
 *  still sees the node and its value — and it belongs to the bucket identity,
 *  not to the key: two `wp:wrapSquare` elements with different `wrapText`
 *  values must not collapse into one verdict. */
function collectPart(
  part: string,
  root: OoxmlNode,
  counts: Map<string, Counted>,
): void {
  const bump = (key: string, inert: boolean) => {
    const id = `${part} ${key} ${inert}`;
    const e = counts.get(id);
    if (e) e.count++;
    else counts.set(id, { part, key, count: 1, inert });
  };
  const walk = (node: OoxmlNode, parentVisited: boolean) => {
    for (const c of node.children) {
      if (subtrees.has(c)) continue; // consumed wholesale by design
      const v = visited.has(c);
      if (v) {
        const asked = readAttrs.get(c);
        for (const a of Object.keys(c.attrs)) {
          // xmlns declarations aren't content — not worth an "ignored" line.
          if (a.startsWith('xmlns')) continue;
          if (!asked?.has(a))
            bump(`${c.name} @${a}`, isInertAttr(c.name, a, c.attrs[a]));
        }
      } else if (parentVisited) {
        bump(c.name, isInertTag(c));
      }
      walk(c, v);
    }
  };
  walk(root, true); // the synthetic #root counts as visited
}

export const audit = {
  /** Whether the audit flag is currently on (re-checked per call, so it can be
   *  toggled live from the console; imports resolve it once in begin). */
  get enabled(): boolean {
    return computeEnabled();
  },

  /** Whether an import is currently COLLECTING — a plain boolean read, unlike
   *  `enabled` which re-resolves the flag through globalThis/localStorage.
   *  Guard extra bookkeeping that only the audit needs with this, so the hot
   *  path pays nothing when the flag is off. */
  get collecting(): boolean {
    return importActive;
  },

  /** Force the flag on/off from code (equivalent to setting the global). */
  setEnabled(on: boolean): void {
    (globalThis as { __BAPBONG_XML_AUDIT__?: boolean }).__BAPBONG_XML_AUDIT__ =
      on;
  },

  /** Register an extra destination for report lines (besides console). */
  setSink(fn: ((line: string) => void) | null): void {
    sink = fn;
  },

  /** The last emitted report (import or export) — programmatic access for
   *  tests and tooling; null until a flagged session completes. */
  get lastReport(): AuditReport | null {
    return lastReport;
  },

  // ── import side ─────────────────────────────────────────────────

  beginImport(label: string): void {
    if (importDepth++ > 0) return; // joined an in-flight session
    importActive = computeEnabled();
    if (!importActive) return;
    importLabel = label;
    visited = new WeakSet();
    subtrees = new WeakSet();
    readAttrs = new WeakMap();
    parts = [];
  },

  /** Hand a parsed part's root to the post-import sweep. */
  registerPart(name: string, root: OoxmlNode): void {
    if (!importActive) return;
    parts.push({ name, root });
  },

  /** Mark an element as touched by the converter. */
  mark(node: OoxmlNode | undefined): void {
    if (!importActive || !node) return;
    visited.add(node);
  },

  /** Mark several elements (the `children()` accessor's result). */
  markAll(nodes: OoxmlNode[]): void {
    if (!importActive) return;
    for (const n of nodes) visited.add(n);
  },

  /** Mark a node and its ENTIRE subtree as consumed (deliberate wholesale
   *  handling: dropped tracked-change content, flattened OMML, …). */
  markSubtree(node: OoxmlNode | undefined): void {
    if (!importActive || !node) return;
    subtrees.add(node);
  },

  /** Record that an attribute was asked for (asked = covered, even if the
   *  attribute is absent in this document). */
  markAttr(node: OoxmlNode | undefined, name: string): void {
    if (!importActive || !node) return;
    let set = readAttrs.get(node);
    if (!set) {
      set = new Set();
      readAttrs.set(node, set);
    }
    set.add(name);
  },

  endImport(): void {
    importDepth = Math.max(0, importDepth - 1);
    if (importDepth > 0 || !importActive) return;
    importActive = false;
    const counts = new Map<string, Counted>();
    for (const p of parts) collectPart(p.name, p.root, counts);
    emitReport({ mode: 'import', label: importLabel, ...classify(counts) });
    parts = [];
  },

  // ── export side ─────────────────────────────────────────────────

  beginExport(label: string): void {
    exportActive = computeEnabled();
    if (!exportActive) return;
    exportLabel = label;
    exportCounts = new Map();
  },

  /** A ProseMirror node/mark type the serializer has no handler for. */
  exportUnhandled(kind: 'node' | 'mark', typeName: string): void {
    if (!exportActive) return;
    const key = `${kind} ${typeName}`;
    exportCounts.set(key, (exportCounts.get(key) ?? 0) + 1);
  },

  endExport(): void {
    if (!exportActive) return;
    exportActive = false;
    const unknown: AuditEntry[] = [...exportCounts.entries()]
      .map(([key, count]) => ({ part: 'model', key, count }))
      .sort((a, b) => b.count - a.count || a.key.localeCompare(b.key));
    emitReport({
      mode: 'export',
      label: exportLabel,
      unknown,
      inert: [],
      ignored: [],
    });
  },
};
