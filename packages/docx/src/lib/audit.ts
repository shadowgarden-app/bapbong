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
 *   - Entries are split into UNKNOWN (real gaps) and ignored-by-design (the
 *     IGNORED_* lists below: revision noise, style-gallery metadata, …).
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

export interface AuditReport {
  mode: 'import' | 'export';
  label: string;
  unknown: AuditEntry[];
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
  // Theme parts beyond colors (and later fonts): effects/object defaults and
  // extra scheme variants style Office's galleries, not document content.
  'a:fmtScheme',
  'a:objectDefaults',
  'a:extraClrSchemeLst',
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
    // Word's visited-link tracking flag — UI state, not content.
    (tag === 'w:hyperlink' && name === 'w:history') ||
    // A bookmark's numeric id pairs start/end within the part; the NAME is
    // what links point at, and the exporter renumbers on the way out.
    (tag === 'w:bookmarkStart' &&
      (name === 'w:id' || name === 'w:displacedByCustomXml')) ||
    // Drawing object ids/names are display metadata; export regenerates them.
    (tag === 'wp:docPr' && (name === 'id' || name === 'name')) ||
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
      ['allowOverlap', 'layoutInCell', 'locked', 'relativeHeight', 'simplePos'].includes(name)) ||
    // CJK disambiguation hint — the explicit rFonts attrs are all read.
    (tag === 'w:rFonts' && name === 'w:hint')
  );
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

function classify(
  counts: Map<string, { part: string; key: string; count: number }>,
): { unknown: AuditEntry[]; ignored: AuditEntry[] } {
  const unknown: AuditEntry[] = [];
  // UNKNOWN entries keep their part (the location matters for fixing them);
  // ignored-by-design ones aggregate across parts — pure noise volume
  // otherwise (the same rsid attr once per header/footer part).
  const ignoredByKey = new Map<string, AuditEntry>();
  for (const e of counts.values()) {
    const tag = e.key.split(' @')[0];
    const attr = e.key.includes(' @') ? e.key.split(' @')[1] : null;
    const isIgnored = attr ? isIgnoredAttr(tag, attr) : isIgnoredTag(tag);
    if (!isIgnored) {
      unknown.push(e);
    } else {
      const agg = ignoredByKey.get(e.key);
      if (agg) agg.count += e.count;
      else ignoredByKey.set(e.key, { part: '*', key: e.key, count: e.count });
    }
  }
  const ignored = [...ignoredByKey.values()];
  const byCount = (a: AuditEntry, b: AuditEntry) =>
    b.count - a.count || a.key.localeCompare(b.key);
  unknown.sort(byCount);
  ignored.sort(byCount);
  return { unknown, ignored };
}

function emitReport(report: AuditReport): void {
  lastReport = report;
  const { mode, label, unknown, ignored } = report;
  const header =
    `[xml-audit] ${mode} "${label}" — ` +
    `${unknown.length} UNKNOWN, ${ignored.length} ignored-by-design`;
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
    ...[...unknown, ...ignored].map((e) => e.part.length),
  );
  for (const e of unknown)
    emitLine(`  UNKNOWN ${e.part.padEnd(pad)}  ${e.key}  ×${e.count}`);
  for (const e of ignored)
    emitLine(`  ignored ${e.part.padEnd(pad)}  ${e.key}  ×${e.count}`);
  if (grouped) c.groupEnd?.();
}

/** Walk a registered part, collecting untouched tags/attrs (see module doc
 *  for the boundary rule). */
function collectPart(
  part: string,
  root: OoxmlNode,
  counts: Map<string, { part: string; key: string; count: number }>,
): void {
  const bump = (key: string) => {
    const id = `${part} ${key}`;
    const e = counts.get(id);
    if (e) e.count++;
    else counts.set(id, { part, key, count: 1 });
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
          if (!asked?.has(a)) bump(`${c.name} @${a}`);
        }
      } else if (parentVisited) {
        bump(c.name);
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
    const counts = new Map<
      string,
      { part: string; key: string; count: number }
    >();
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
    emitReport({ mode: 'export', label: exportLabel, unknown, ignored: [] });
  },
};
