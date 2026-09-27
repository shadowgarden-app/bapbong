import type JSZip from 'jszip';
import { attrOf, child, children, parseXml } from './ooxml.js';

/**
 * Package parts a verbatim-carried drawing depends on.
 *
 * A chart's drawing XML says only `<c:chart r:id="rId12"/>`: the chart itself
 * lives in its own part (word/charts/chart1.xml), which in turn relates to an
 * embedded workbook, a theme override, style and colour parts. Carrying the
 * drawing alone would write a reference to nothing — Word calls that file
 * damaged — so the drawing travels with the parts it reaches, and the
 * exporter writes them back (or keeps the originals, when the saved package
 * still holds exactly these).
 *
 * Plain data (it rides a node attr): paths are package-absolute, without a
 * leading slash ("word/charts/chart1.xml").
 */

/** One relationship, target resolved. Internal targets are package paths;
 *  external ones stay verbatim (a URL). */
export interface CarriedRel {
  id: string;
  type: string;
  target: string;
  external?: boolean;
}

/** One part: its bytes (text for XML parts, base64 otherwise), its content
 *  type, and its own relationships. */
export interface CarriedPart {
  path: string;
  contentType: string;
  data: string;
  base64?: boolean;
  rels?: CarriedRel[];
}

/** The package's `[Content_Types].xml`: per-part overrides, then defaults
 *  by extension. */
export interface ContentTypes {
  overrides: Map<string, string>;
  defaults: Map<string, string>;
}

export function parseContentTypes(xml: string | undefined): ContentTypes {
  const overrides = new Map<string, string>();
  const defaults = new Map<string, string>();
  const types = child(xml ? parseXml(xml) : undefined, 'Types');
  for (const o of children(types, 'Override')) {
    const name = attrOf(o, 'PartName');
    const ct = attrOf(o, 'ContentType');
    if (name && ct) overrides.set(name.replace(/^\/+/, ''), ct);
  }
  for (const d of children(types, 'Default')) {
    const ext = attrOf(d, 'Extension');
    const ct = attrOf(d, 'ContentType');
    if (ext && ct) defaults.set(ext.toLowerCase(), ct);
  }
  return { overrides, defaults };
}

function contentTypeOf(path: string, types: ContentTypes): string | undefined {
  const override = types.overrides.get(path);
  if (override) return override;
  const dot = path.lastIndexOf('.');
  return dot < 0
    ? undefined
    : types.defaults.get(path.slice(dot + 1).toLowerCase());
}

/** A relationship target resolved against the part that holds it: relative
 *  to that part's folder, or to the package root when it starts with "/"
 *  (OPC §9.3). "word/charts/chart1.xml" + "../embeddings/a.xlsx" →
 *  "word/embeddings/a.xlsx". */
export function resolveTarget(sourcePart: string, target: string): string {
  const parts = target.startsWith('/')
    ? []
    : sourcePart.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') parts.pop();
    else parts.push(seg);
  }
  return parts.join('/');
}

/** The inverse: the target a part at `sourcePart` writes to reach `path`. */
export function relativeTarget(sourcePart: string, path: string): string {
  const from = sourcePart.split('/').slice(0, -1);
  const to = path.split('/');
  let common = 0;
  while (
    common < from.length &&
    common < to.length - 1 &&
    from[common] === to[common]
  )
    common++;
  return [
    ...Array<string>(from.length - common).fill('..'),
    ...to.slice(common),
  ].join('/');
}

/** "word/charts/chart1.xml" → "word/charts/_rels/chart1.xml.rels". */
export function relsPathOf(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  return `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
}

/** Content types the package writes as text; everything else is bytes. */
function isXml(contentType: string, path: string): boolean {
  return /[+/]xml$/.test(contentType) || path.toLowerCase().endsWith('.xml');
}

/** The relationships a part declares, targets resolved. */
export function parseRels(
  partPath: string,
  xml: string | undefined,
): CarriedRel[] {
  const out: CarriedRel[] = [];
  const root = child(xml ? parseXml(xml) : undefined, 'Relationships');
  for (const rel of children(root, 'Relationship')) {
    const id = attrOf(rel, 'Id');
    const type = attrOf(rel, 'Type');
    const target = attrOf(rel, 'Target');
    if (!id || !type || !target) continue;
    const external = attrOf(rel, 'TargetMode') === 'External';
    out.push(
      external
        ? { id, type, target, external }
        : { id, type, target: resolveTarget(partPath, target) },
    );
  }
  return out;
}

const esc = (s: string) =>
  s.replace(
    /[&<>"]/g,
    (c) =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c] as string,
  );

const PR_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** A `.rels` part for relationships whose targets are already relative. */
export function relationshipsXml(rels: CarriedRel[]): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PR_NS}">` +
    rels
      .map(
        (r) =>
          `<Relationship Id="${esc(r.id)}" Type="${esc(r.type)}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`,
      )
      .join('') +
    '</Relationships>'
  );
}

/** Parts the exporter writes beside the regenerated body, shared by every
 *  story (body, header and footer parts). */
export interface PartSink {
  /** Paths the output package already holds, lower-cased: OPC part names
   *  compare case-insensitively (§9.1.1.1). */
  taken: Set<string>;
  files: { path: string; data: string; base64?: boolean }[];
  /** `[Content_Types].xml` overrides for the parts written. */
  overrides: string[];
}

export function partSink(existing: Iterable<string> = []): PartSink {
  return {
    taken: new Set([...existing].map((p) => p.toLowerCase())),
    files: [],
    overrides: [],
  };
}

/** A free name beside `path`: same folder and extension, the stem's trailing
 *  number replaced by the first one not taken ("chart1.xml" → "chart6.xml"). */
function freshPath(path: string, taken: Set<string>): string {
  const slash = path.lastIndexOf('/');
  const dir = path.slice(0, slash + 1);
  const name = path.slice(slash + 1);
  const dot = name.lastIndexOf('.');
  const ext = dot > 0 ? name.slice(dot) : '';
  const stem = (dot > 0 ? name.slice(0, dot) : name).replace(/\d+$/, '');
  for (let n = 1; ; n++) {
    const candidate = `${dir}${stem}${n}${ext}`;
    if (!taken.has(candidate.toLowerCase())) {
      taken.add(candidate.toLowerCase());
      return candidate;
    }
  }
}

/**
 * Write `root` and every part it reaches under fresh names, relationships
 * re-pointed at the new names. Returns the root's new path — null when
 * `root` is not among `parts`. A dependency the import could not load keeps
 * its original target, as dangling as it was in the source.
 */
export function writePartClosure(
  root: string,
  parts: CarriedPart[],
  sink: PartSink,
): string | null {
  const byPath = new Map(parts.map((p) => [p.path, p]));
  const renamed = new Map<string, string>();
  const order: CarriedPart[] = [];
  const walk = (path: string): void => {
    const part = byPath.get(path);
    if (!part || renamed.has(path)) return;
    renamed.set(path, freshPath(path, sink.taken));
    order.push(part);
    for (const rel of part.rels ?? []) if (!rel.external) walk(rel.target);
  };
  walk(root);
  for (const part of order) {
    const at = renamed.get(part.path) as string;
    sink.files.push({
      path: at,
      data: part.data,
      ...(part.base64 && { base64: true }),
    });
    sink.overrides.push(
      `<Override PartName="/${esc(at)}" ContentType="${esc(part.contentType)}"/>`,
    );
    if (part.rels?.length)
      sink.files.push({
        path: relsPathOf(at),
        data: relationshipsXml(
          part.rels.map((r) =>
            r.external
              ? r
              : {
                  ...r,
                  target: relativeTarget(at, renamed.get(r.target) ?? r.target),
                },
          ),
        ),
      });
  }
  return renamed.get(root) ?? null;
}

/**
 * Does `carry` hold `root`'s closure exactly as `parts` has it — every part
 * at the same path with the same bytes and the same relationships? Then the
 * saved package can keep the original parts and relationship instead of
 * writing a copy.
 */
export async function carriedVerbatim(
  root: string,
  parts: CarriedPart[],
  carry: JSZip,
): Promise<boolean> {
  const byPath = new Map(parts.map((p) => [p.path, p]));
  const seen = new Set<string>();
  const same = async (path: string): Promise<boolean> => {
    if (seen.has(path)) return true;
    seen.add(path);
    const part = byPath.get(path);
    const entry = carry.file(path);
    if (!part || !entry || entry.dir) return false;
    if ((await entry.async(part.base64 ? 'base64' : 'string')) !== part.data)
      return false;
    const key = (rels: CarriedRel[]) =>
      JSON.stringify(
        [...rels]
          .sort((a, b) => (a.id < b.id ? -1 : 1))
          .map((r) => [r.id, r.type, r.target, !!r.external]),
      );
    const carried = parseRels(
      path,
      await carry.file(relsPathOf(path))?.async('string'),
    );
    if (key(carried) !== key(part.rels ?? [])) return false;
    for (const rel of part.rels ?? [])
      if (!rel.external && byPath.has(rel.target) && !(await same(rel.target)))
        return false;
    return true;
  };
  return same(root);
}

/**
 * `rootPath` and every part it reaches through relationships, root first.
 * Null when the root is missing or has no known content type — carrying it
 * would write a part Word cannot place. A missing DEPENDENCY is left out
 * with its relationship kept as the source wrote it: Word opens the source
 * that way, so the carry is no worse.
 */
export async function loadPartClosure(
  zip: JSZip,
  rootPath: string,
  types: ContentTypes,
): Promise<CarriedPart[] | null> {
  const out: CarriedPart[] = [];
  const seen = new Set<string>();
  const visit = async (path: string): Promise<boolean> => {
    if (seen.has(path)) return true;
    seen.add(path);
    const entry = zip.file(path);
    const contentType = contentTypeOf(path, types);
    if (!entry || entry.dir || !contentType) return false;
    const base64 = !isXml(contentType, path);
    const part: CarriedPart = {
      path,
      contentType,
      data: await entry.async(base64 ? 'base64' : 'string'),
      ...(base64 && { base64 }),
    };
    out.push(part);
    const relsXml = await zip.file(relsPathOf(path))?.async('string');
    const rels = parseRels(path, relsXml);
    if (rels.length > 0) part.rels = rels;
    for (const rel of rels) if (!rel.external) await visit(rel.target);
    return true;
  };
  return (await visit(rootPath)) ? out : null;
}
