import type { ResolvedTableStyle } from '@shadow-garden/bapbong-contracts';
import { parseXml } from './ooxml.js';
import { buildStyleRegistry } from './styles.js';
import { buildTableStyleSheet } from './docx.js';

/**
 * The built-in table-style catalog: what the gallery offers a document whose
 * own sheet doesn't already carry the style. Thirteen styles — Table Grid,
 * Light Grid and Medium Shading 1 across Word's six accent colours — with
 * WORD'S OWN styleIds and names, so a document that already defines one
 * (imported) simply wins by id, and Word's gallery recognises what we wrote.
 *
 * One source of truth: each entry is OOXML. The resolved form the gallery
 * and renderer use comes from running that XML through the same
 * buildStyleRegistry → buildTableStyleSheet pipeline imports use, and the
 * exporter appends the XML itself into styles.xml — so the preview, the
 * page, and what Word opens are three readings of one definition. The shape
 * is the probe-F2 style, which Word's own PDF verified cell by cell.
 *
 * Deliberately self-contained (no basedOn, cell margins inline): a catalog
 * style must land in ANY styles.xml — including a from-scratch one whose
 * TableNormal we don't control — without a dangling reference.
 */

/** Word's theme accent colours (Office theme), literal — a catalog style
 *  must not change colour with the document's theme part. */
const ACCENTS: ReadonlyArray<readonly [number, string]> = [
  [1, '4F81BD'],
  [2, 'C0504D'],
  [3, '9BBB59'],
  [4, '8064A2'],
  [5, '4BACC6'],
  [6, 'F79646'],
];

/** `f` of the colour over white. 0.405 is calibrated against the ONE band
 *  literal we hold Word's own bytes for (accent3 → D6E3BC, from the Kpop
 *  corpus file, all three channels); the other accents get the same mix —
 *  close to Word's gallery, and exactly what our own XML will render as,
 *  which is the consistency that actually matters. */
function tint(hex: string, f: number): string {
  const n = parseInt(hex, 16);
  const ch = (s: number) =>
    Math.round(((n >> s) & 255) * f + 255 * (1 - f))
      .toString(16)
      .toUpperCase()
      .padStart(2, '0');
  return `${ch(16)}${ch(8)}${ch(0)}`;
}

const side = (s: string, color: string, sz = 8, val = 'single') =>
  `<w:${s} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${color}"/>`;
const CELL_MAR =
  '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>' +
  '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>';
const BANDS =
  '<w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/>';
const NO_SPACING =
  '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>';

function lightGrid(n: number, accent: string): string {
  const band = tint(accent, 0.405);
  const all = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((s) => side(s, accent))
    .join('');
  const boxV = ['top', 'left', 'bottom', 'right', 'insideV']
    .map((s) => side(s, accent))
    .join('');
  return (
    `<w:style w:type="table" w:styleId="LightGrid-Accent${n}"><w:name w:val="Light Grid Accent ${n}"/><w:uiPriority w:val="62"/>` +
    `${NO_SPACING}<w:tblPr>${BANDS}<w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>${all}</w:tblBorders>${CELL_MAR}</w:tblPr>` +
    `<w:tblStylePr w:type="firstRow">${NO_SPACING}<w:rPr><w:b/></w:rPr><w:tcPr><w:tcBorders>` +
    side('top', accent) +
    side('left', accent) +
    side('bottom', accent, 18) +
    side('right', accent) +
    side('insideH', 'auto', 0, 'nil') +
    side('insideV', accent) +
    `</w:tcBorders></w:tcPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="lastRow">${NO_SPACING}<w:rPr><w:b/></w:rPr><w:tcPr><w:tcBorders>` +
    side('top', accent, 6, 'double') +
    side('left', accent) +
    side('bottom', accent) +
    side('right', accent) +
    side('insideH', 'auto', 0, 'nil') +
    side('insideV', accent) +
    `</w:tcBorders></w:tcPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="firstCol"><w:rPr><w:b/></w:rPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="lastCol"><w:rPr><w:b/></w:rPr><w:tcPr><w:tcBorders>${side('left', accent, 18)}</w:tcBorders></w:tcPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="band1Vert"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${band}"/></w:tcPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="band1Horz"><w:tcPr><w:tcBorders>${boxV}</w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="${band}"/></w:tcPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="band2Horz"><w:tcPr><w:tcBorders>${boxV}</w:tcBorders></w:tcPr></w:tblStylePr>` +
    `</w:style>`
  );
}

function mediumShading(n: number, accent: string): string {
  const border = tint(accent, 0.75);
  const band = tint(accent, 0.405);
  const outer = ['top', 'left', 'bottom', 'right', 'insideH']
    .map((s) => side(s, border))
    .join('');
  return (
    `<w:style w:type="table" w:styleId="MediumShading1-Accent${n}"><w:name w:val="Medium Shading 1 Accent ${n}"/><w:uiPriority w:val="63"/>` +
    `${NO_SPACING}<w:tblPr>${BANDS}<w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>${outer}</w:tblBorders>${CELL_MAR}</w:tblPr>` +
    `<w:tblStylePr w:type="firstRow">${NO_SPACING}<w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:tcPr><w:tcBorders>` +
    side('top', border) +
    side('left', border) +
    side('bottom', border) +
    side('right', border) +
    side('insideH', 'auto', 0, 'nil') +
    side('insideV', 'auto', 0, 'nil') +
    `</w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="${accent}"/></w:tcPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="lastRow">${NO_SPACING}<w:rPr><w:b/></w:rPr><w:tcPr><w:tcBorders>` +
    side('top', border, 6, 'double') +
    `</w:tcBorders></w:tcPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="firstCol"><w:rPr><w:b/></w:rPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="lastCol"><w:rPr><w:b/></w:rPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${band}"/></w:tcPr></w:tblStylePr>` +
    `</w:style>`
  );
}

const TABLE_GRID =
  `<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:uiPriority w:val="39"/>` +
  `${NO_SPACING}<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>` +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((s) => side(s, '000000', 4))
    .join('') +
  `</w:tblBorders>${CELL_MAR}</w:tblPr></w:style>`;

export interface TableStyleCatalogEntry {
  id: string;
  /** Word's gallery name ("Light Grid Accent 3"). */
  name: string;
  /** The `<w:style>` element, appended to styles.xml on save when the
   *  document doesn't define the id. */
  xml: string;
}

export const TABLE_STYLE_CATALOG: readonly TableStyleCatalogEntry[] = [
  { id: 'TableGrid', name: 'Table Grid', xml: TABLE_GRID },
  ...ACCENTS.map(([n, hex]) => ({
    id: `LightGrid-Accent${n}`,
    name: `Light Grid Accent ${n}`,
    xml: lightGrid(n, hex),
  })),
  ...ACCENTS.map(([n, hex]) => ({
    id: `MediumShading1-Accent${n}`,
    name: `Medium Shading 1 Accent ${n}`,
    xml: mediumShading(n, hex),
  })),
];

/** The catalog entry's XML, for the exporter's styles.xml merge. */
export function catalogStyleXml(id: string): string | undefined {
  return TABLE_STYLE_CATALOG.find((e) => e.id === id)?.xml;
}

let resolved: ReadonlyArray<{
  id: string;
  name: string;
  style: ResolvedTableStyle;
}> | null = null;

/** The catalog in the gallery/renderer's form, resolved ONCE through the
 *  exact pipeline imports use — the preview and the saved file cannot
 *  diverge because both read this XML. */
export function catalogTableStyles(): ReadonlyArray<{
  id: string;
  name: string;
  style: ResolvedTableStyle;
}> {
  if (resolved) return resolved;
  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const xml =
    `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">` +
    TABLE_STYLE_CATALOG.map((e) => e.xml).join('') +
    `</w:styles>`;
  const registry = buildStyleRegistry(parseXml(xml));
  const sheet = buildTableStyleSheet(
    TABLE_STYLE_CATALOG.map((e) => e.id),
    registry,
  );
  resolved = TABLE_STYLE_CATALOG.map((e) => ({
    id: e.id,
    name: e.name,
    style: sheet[e.id],
  }));
  return resolved;
}
