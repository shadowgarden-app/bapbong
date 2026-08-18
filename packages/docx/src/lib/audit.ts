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
  /** Inside a declaration nothing in the document points at — see
   *  {@link collectPart}'s dead-zone walk. */
  unreferenced: boolean;
}

export interface AuditReport {
  mode: 'import' | 'export';
  label: string;
  unknown: AuditEntry[];
  /** Unread, but carrying a value the spec defines as having no effect — so
   *  reading it would change nothing on the page. See INERT_ATTRS. */
  inert: AuditEntry[];
  /**
   * Unread, and inside something the document DECLARES but never uses: a
   * `w:style` whose id is mentioned nowhere, or the theme's `a:fmtScheme`
   * when no shape references it. Nothing on the page comes from these, so
   * counting them as gaps buries the real ones — but they are a separate
   * bucket from `inert`, because the reason is reachability rather than
   * value, and from `ignored`, because we did not decide to skip them.
   */
  unreferenced: AuditEntry[];
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
  // The document's list of smart-tag recognisers. Same subject as the
  // w:smartTag attrs below: the tags themselves unwrap to their runs.
  'w:smartTagType',
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
  // whole settings.xml part is carried on export. (w:compat is NOT here: it is
  // read — see compat.ts — and its children are judged one by one below.)
  'w:decimalSymbol',
  'w:characterSpacingControl',
  'w:hideSpellingErrors',
  'w:hideGrammaticalErrors',
  'w:displayBackgroundShape',
  'w:formProt',
  // Same family: the alignment grid Word draws in its own window, whether it
  // recompresses images on save, whether subdocuments count towards the word
  // count, and a legacy-interop toggle. None of them reaches the page.
  'w:drawingGridHorizontalSpacing',
  'w:drawingGridVerticalSpacing',
  'w:displayHorizontalDrawingGridEvery',
  'w:displayVerticalDrawingGridEvery',
  'w:doNotAutoCompressPictures',
  'w:doNotIncludeSubdocsInStats',
  'w:uiCompat97To2003',
  // ── w:compat children not adopted into DocCompat ─────────────────────
  // Every one of these is a real Word layout switch, and each is a candidate
  // field for DocCompat (compat.ts) the day its rule is implemented. Listed by
  // NAME on purpose: a flag Word writes that is not here still surfaces as
  // UNKNOWN, which is the signal to go and read up on it. The two compat
  // children that ARE consulted (compatSetting/compatibilityMode and
  // doNotUseHTMLParagraphAutoSpacing) are read and never reach this list.
  //
  // Word 2003→2007 conversion set (a converted .doc carries all of these):
  'w:useNormalStyleForList',
  'w:doNotUseIndentAsNumberingTabStop',
  'w:useAltKinsokuLineBreakRules',
  'w:allowSpaceOfSameStyleInTable',
  'w:doNotSuppressIndentation',
  'w:doNotAutofitConstrainedTables',
  'w:autofitToFirstFixedWidthCell',
  'w:underlineTabInNumList',
  'w:displayHangulFixedWidth',
  'w:splitPgBreakAndParaMark',
  'w:doNotVertAlignCellWithSp',
  'w:doNotBreakConstrainedForcedTable',
  'w:doNotVertAlignInTxbx',
  'w:useAnsiKerningPairs',
  'w:cachedColBalance',
  // East-Asian layout switches — no CJK line-breaking or width balancing yet.
  'w:useFELayout',
  'w:balanceSingleByteDoubleByteWidth',
  'w:doNotLeaveBackslashAlone',
  // (w:ulTrailSpace and w:doNotExpandShiftReturn are READ — DocCompat.)
  // Word's own drawing grid origin — UI, like the other drawingGrid* above.
  'w:doNotUseMarginsForDrawingGridOrigin',
  // An embedded object's link to its editor: ProgID, the OLE stream, the
  // shape it belongs to. What the reader SEES is the v:imagedata preview
  // beside it, which is read and rendered. Editing the embedded object is
  // not something this program does, so the link has nothing to drive.
  'o:OLEObject',
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

/**
 * `w:compatSetting` entries Word 2010+ writes on every document that are not
 * adopted into DocCompat — the report keys them `w:compatSetting[<name>]` so
 * each is judged on its own (see {@link keyOf}). Same contract as the element
 * list above: a setting Word starts writing that is not here surfaces as
 * UNKNOWN. compatibilityMode and overrideTableStyleFontSizeAndJustification
 * are READ (compat.ts) and never reach this list.
 */
const COMPAT_SETTINGS_NOT_ADOPTED = new Set([
  // OpenType typographic features (ligatures, stylistic sets): the per-run
  // w14:* requests that would use them are ignored by prefix already, so the
  // master switch has nothing to enable.
  'enableOpenTypeFeatures',
  // Mirror indents on facing pages: w:mirrorIndents itself is not modelled
  // and stays countable wherever a paragraph uses it; the flip rule is moot.
  'doNotFlipMirrorIndents',
  // Hyphenation is out of layout scope by decision (see w:autoHyphenation).
  'useWord2013TrackBottomHyphenation',
  'allowHyphenationAtTrackBottom',
  // Repeated header rows: whether Word treats a multi-row header block as
  // one; the rows repeat either way and nothing here reads the difference.
  'differentiateMultirowTableHeaders',
  // Floating (tblpPr) tables are not modelled — the break rule has no table
  // to apply to.
  'allowTextAfterFloatingTableBreak',
]);

/** The report key of an element: its tag, except a `w:compatSetting`, which
 *  is `w:compatSetting[<w:name>]` — the name IS the setting's identity, and
 *  one entry per setting is what makes the compat list above readable. */
function keyOf(node: OoxmlNode): string {
  return node.name === 'w:compatSetting'
    ? `w:compatSetting[${node.attrs['w:name'] ?? ''}]`
    : node.name;
}

/** Attribute names skipped on purpose (on otherwise-covered elements). */
function isIgnoredAttr(tag: string, name: string): boolean {
  if (tag.startsWith('w:compatSetting')) {
    // The namespace URI is the same constant on every entry; the name is
    // always read (it is how an entry is found) and rides the key.
    if (name === 'w:uri' || name === 'w:name') return true;
    const setting = tag.slice('w:compatSetting['.length, -1);
    if (name === 'w:val' && COMPAT_SETTINGS_NOT_ADOPTED.has(setting))
      return true;
  }
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
    // The printer's paper code (ST_DecimalNumber naming a tray/form). Page
    // geometry comes from @w:w/@w:h/@w:orient, which are read; this one
    // addresses a print driver we do not talk to.
    (tag === 'w:pgSz' && name === 'w:code') ||
    // The OLE siblings of o:OLEObject above: the flag marking a VML shape as
    // an embedded object, and the object's ORIGINAL size — the shape's own
    // style width/height is what Word displays and what we read.
    (tag.startsWith('v:') && name === 'o:ole') ||
    (tag === 'w:object' && (name === 'w:dxaOrig' || name === 'w:dyaOrig')) ||
    // Whether a style came from Word's built-in gallery or the user made it:
    // drives the Styles pane and built-in name mapping, never rendering.
    (tag === 'w:style' && name === 'w:customStyle') ||
    // Word's visited-link tracking flag — UI state, not content.
    (tag === 'w:hyperlink' && name === 'w:history') ||
    // Which recogniser claimed the text ("place", "country-region", and the
    // namespace it came from). The tag is unwrapped to its runs and never
    // re-emitted, so the recogniser's identity has nothing to act on.
    (tag === 'w:smartTag' && (name === 'w:element' || name === 'w:uri')) ||
    (tag === 'w:customXml' && (name === 'w:element' || name === 'w:uri')) ||
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
/** A wrap polygon's coordinate space: 0..21600 of the shape's own extent. */
const WRAP_POLY_FULL = 21600;
const isFalse = (v: string) => FALSE_VALUES.has(v);
const isNum = (n: number) => (v: string) => Number(v) === n;

/** `tag @attr` → predicate on the raw value AND the element it sits on: true
 *  = this value is a no-op. The node is there for the handful of attributes
 *  whose effect depends on a sibling attribute (`@upright` only matters to a
 *  rotated shape); most predicates ignore it. */
const INERT_ATTRS: Record<string, (v: string, n: OoxmlNode) => boolean> = {
  // CT_TextBodyProperties. Word stamps the whole default set onto every
  // textbox it writes. NOT listed: @anchor, which is READ now (the text block
  // slides to ctr/b), and @compatLnSpc, whose default is false — a written
  // "1" really does change line spacing inside the shape.
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
  'wps:bodyPr @anchorCtr': isFalse,
  'wps:bodyPr @forceAA': isFalse,
  // "Keep the text upright while the shape is rotated" — nothing to keep
  // upright when the shape is not rotated. @rot sits on this same element.
  'wps:bodyPr @upright': (_v, n) => Number(n.attrs['rot'] ?? '0') === 0,
  // CT_TblWidth: with @w:type="nil" there is no width, so the @w:w beside it
  // says nothing (the reader takes 0 from the type alone). Any other type
  // reads the value, and an unread one is a real gap.
  'w:tblInd @w:w': (_v, n) => n.attrs['w:type'] === 'nil',
  // ST_WrapText defaults to bothSides: text flows down both sides of the
  // float, which is what our square-wrap already does. left/right/largest
  // pick ONE side and stay UNKNOWN.
  'wp:wrapSquare @wrapText': (v) => v === 'bothSides',
  // The same attribute on the same terms — tight and through wrap import as
  // square, so which sides they use is decided the same way.
  'wp:wrapTight @wrapText': (v) => v === 'bothSides',
  'wp:wrapThrough @wrapText': (v) => v === 'bothSides',
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
  // Automatic spacing between East-Asian and Latin/numeric text, turned OFF.
  // We never insert that spacing, so "off" asks for exactly what we do. ON —
  // including the absent-means-on default — stays UNKNOWN: it is a real
  // feature we do not have, and a file that wants it must say so in the
  // report.
  'w:autoSpaceDE': (n) => isFalse(String(n.attrs['w:val'] ?? '')),
  'w:autoSpaceDN': (n) => isFalse(String(n.attrs['w:val'] ?? '')),
  // A wrap polygon that IS the shape's box: Word writes one on every
  // tight/through-wrapped picture whose outline it never traced. The
  // coordinate space is 0..21600 of the extent, and Word rounds the far edge
  // a few thousandths short (21435 of 21600 in the factsheet), so the corners
  // are compared with a tolerance. Anything that actually carves a shape —
  // more than four corners, or a rectangle inset from the box, which would
  // let text closer than our rectangle does — stays UNKNOWN, because wrapping
  // to a contour is a feature we do not have.
  'wp:wrapPolygon': (n) => {
    const pts = n.children
      .filter((c) => c.name === 'wp:start' || c.name === 'wp:lineTo')
      .map((c) => [Number(c.attrs['x']), Number(c.attrs['y'])] as const);
    if (pts.length < 4 || pts.some(([x, y]) => !isFinite(x) || !isFinite(y)))
      return false;
    const spansBox = (vals: number[]) => {
      const uniq = [...new Set(vals)];
      return (
        uniq.length === 2 &&
        Math.min(...uniq) === 0 &&
        Math.max(...uniq) >= WRAP_POLY_FULL * 0.98
      );
    };
    return spansBox(pts.map((p) => p[0])) && spansBox(pts.map((p) => p[1]));
  },
  // ST_DocGrid @w:type defaults to "default", which snaps nothing — and that
  // makes the linePitch/charSpace it carries moot. Only lines/linesAndChars
  // actually grid the page, and those stay UNKNOWN.
  'w:docGrid': (n) => {
    const t = n.attrs['w:type'];
    return t === undefined || t === 'default';
  },
  // settings.xml's footnote/endnote block. Word writes it on every document
  // just to point at the two special notes that hold the separator rule and
  // its continuation (w:id="-1" and "0"). In that shape it configures
  // nothing: the properties that would — w:numFmt, w:pos, w:numRestart,
  // w:numStart — are absent, and any of them present drops it back to
  // UNKNOWN.
  'w:footnotePr': (n) => n.children.every((c) => c.name === 'w:footnote'),
  'w:endnotePr': (n) => n.children.every((c) => c.name === 'w:endnote'),
  // A content control's chrome. The control itself is unwrapped to its
  // content, and Word copies the control's rPr onto the runs inside it — so
  // an sdtPr holding only ids and gallery metadata has nothing left to say.
  // One carrying its OWN w:rPr might (we would be dropping formatting), so
  // that shape stays UNKNOWN.
  'w:sdtPr': (n) => !n.children.some((c) => c.name === 'w:rPr'),
  'w:sdtEndPr': (n) => !n.children.some((c) => c.name === 'w:rPr'),
  // A conditional table format Word itself throws away: "Word does not apply
  // and discards on save any properties within the tblStylePr element when
  // the type attribute has a value of wholeTable" (MS-OI29500 §17.18.89(a)).
  // Every other @w:type IS read (styles.ts resolveCond), so this stays a
  // value-based demotion — a firstRow branch we failed to handle would still
  // be reported. Whole-table formatting reaches content through the style's
  // own w:pPr/w:rPr/w:tblPr instead.
  'w:tblStylePr': (n) => n.attrs['w:type'] === 'wholeTable',
  // An element with no children and no attributes says nothing. Word writes
  // `<w:tblPr/>` inside most conditional branches of its built-in table
  // styles, and it would have little to say even if it were filled: "Word does
  // not allow [bidiVisual, tblLayout, tblLook, tblOverlap, tblpPr, tblStyle,
  // tblStyleColBandSize, tblStyleRowBandSize, tblW] to be child elements of
  // the tblPr element" there (MS-OI29500 §17.7.6.3). A non-empty one still
  // reports.
  'w:tblPr': (n) => n.children.length === 0 && attrCount(n) === 0,
};

function attrCount(n: OoxmlNode): number {
  return Object.keys(n.attrs).filter((a) => !a.startsWith('xmlns')).length;
}

function isInertAttr(
  tag: string,
  name: string,
  value: string,
  node: OoxmlNode,
): boolean {
  return INERT_ATTRS[`${tag} @${name}`]?.(value, node) ?? false;
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
/** Whether any shape points into the theme's format scheme — the input to the
 *  dead-zone rule. Filled by scanThemeReferences at endImport. */
let themeReferenced = false;
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
  unreferenced: AuditEntry[];
  ignored: AuditEntry[];
} {
  const unknown: AuditEntry[] = [];
  // UNKNOWN entries keep their part (the location matters for fixing them);
  // inert and ignored-by-design ones aggregate across parts — pure noise
  // volume otherwise (the same rsid attr once per header/footer part).
  const inertByKey = new Map<string, AuditEntry>();
  const deadByKey = new Map<string, AuditEntry>();
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
    } else if (e.unreferenced) {
      aggregate(deadByKey, e);
    } else {
      unknown.push({ part: e.part, key: e.key, count: e.count });
    }
  }
  const byCount = (a: AuditEntry, b: AuditEntry) =>
    b.count - a.count || a.key.localeCompare(b.key);
  unknown.sort(byCount);
  const inert = [...inertByKey.values()].sort(byCount);
  const unreferenced = [...deadByKey.values()].sort(byCount);
  const ignored = [...ignoredByKey.values()].sort(byCount);
  return { unknown, inert, unreferenced, ignored };
}

function emitReport(report: AuditReport): void {
  lastReport = report;
  const { mode, label, unknown, inert, ignored } = report;
  const unref = report.unreferenced ?? [];
  const header =
    `[xml-audit] ${mode} "${label}" — ` +
    `${unknown.length} UNKNOWN, ${inert.length} inert, ` +
    `${unref.length} unreferenced, ${ignored.length} ignored-by-design`;
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
    ...[...unknown, ...inert, ...unref, ...ignored].map((e) => e.part.length),
  );
  for (const e of unknown)
    emitLine(`  UNKNOWN ${e.part.padEnd(pad)}  ${e.key}  ×${e.count}`);
  for (const e of inert)
    emitLine(`  inert   ${e.part.padEnd(pad)}  ${e.key}  ×${e.count}`);
  for (const e of unref)
    emitLine(`  unref   ${e.part.padEnd(pad)}  ${e.key}  ×${e.count}`);
  for (const e of ignored)
    emitLine(`  ignored ${e.part.padEnd(pad)}  ${e.key}  ×${e.count}`);
  if (grouped) c.groupEnd?.();
}

/**
 * Does any shape in the package point into the theme's style matrix?
 *
 * `a:fmtScheme` is a menu of fills, lines and effects that shapes select from
 * by index — `a:lnRef idx="2"` means "the second entry of lnStyleLst". A
 * document with no such reference renders nothing from it, however many
 * gradients it declares. (`a:clrScheme` is deliberately NOT part of this: theme
 * colours are resolved by name, and stay accountable.)
 *
 * Note what this does NOT do: styles. Unreferenced `w:style` elements are
 * already handled at the source — `auditMarkUnusedStyles` in styles.ts marks
 * their subtrees consumed, with the same "a w:default style is never unused"
 * guard. An earlier cut of this duplicated that rule here and the test proved
 * it dead: an orphan style's properties never reach the sweep at all.
 */
function scanThemeReferences(): void {
  themeReferenced = false;
  const walk = (n: OoxmlNode): void => {
    if (
      n.name === 'a:lnRef' ||
      n.name === 'a:fillRef' ||
      n.name === 'a:effectRef' ||
      n.name === 'a:fontRef'
    ) {
      themeReferenced = true;
      return;
    }
    for (const c of n.children) walk(c);
  };
  for (const p of parts) {
    walk(p.root);
    if (themeReferenced) return;
  }
}

/** Does this node open a declaration nothing in the document points at? */
function isDeadDeclaration(n: OoxmlNode): boolean {
  return n.name === 'a:fmtScheme' && !themeReferenced;
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
  const bump = (key: string, inert: boolean, dead: boolean) => {
    const id = `${part} ${key} ${inert} ${dead}`;
    const e = counts.get(id);
    if (e) e.count++;
    else counts.set(id, { part, key, count: 1, inert, unreferenced: dead });
  };
  const walk = (node: OoxmlNode, parentVisited: boolean, dead: boolean) => {
    for (const c of node.children) {
      if (subtrees.has(c)) continue; // consumed wholesale by design
      // A declaration nothing points at makes a dead zone for its subtree.
      const inDead = dead || isDeadDeclaration(c);
      const v = visited.has(c);
      if (v) {
        const asked = readAttrs.get(c);
        for (const a of Object.keys(c.attrs)) {
          // xmlns declarations aren't content — not worth an "ignored" line.
          if (a.startsWith('xmlns')) continue;
          if (!asked?.has(a))
            bump(
              `${keyOf(c)} @${a}`,
              isInertAttr(c.name, a, c.attrs[a], c),
              inDead,
            );
        }
      } else if (parentVisited) {
        bump(keyOf(c), isInertTag(c), inDead);
      }
      walk(c, v, inDead);
    }
  };
  walk(root, true, false); // the synthetic #root counts as visited
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
    scanThemeReferences();
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
      unreferenced: [],
      ignored: [],
    });
  },
};
