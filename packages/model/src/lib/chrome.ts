/**
 * Which header and footer stories each section actually shows.
 *
 * The importer hands back the stories it read (`sectionChrome`, one entry per
 * section with Word's "Link to Previous" already resolved, or only the flat
 * set when every section shares it), and an edit made in bapbong rides the
 * document as `doc.attrs.sectionChromeOverrides` (a story per section, variant
 * and part, as ProseMirror JSON). The layout, the editor and an agent reading
 * the document all need the two combined the same way — this is that one
 * place. Generic over the story type so it stays free of any schema.
 */

/** One section's stories, by variant ("default" | "first" | "even"). */
export interface ChromeStories<N> {
  headers: Record<string, N>;
  footers: Record<string, N>;
  titlePg: boolean;
}

/** What the importer read, as the view and the headless session keep it. */
export interface ImportedChrome<N> {
  /** Per section, when the file gives sections their own chrome. */
  sectionChrome: ChromeStories<N>[] | null;
  /** The flat set every section shares otherwise. */
  headers: Record<string, N>;
  footers: Record<string, N>;
  titlePg: boolean;
}

/**
 * The per-section chrome in effect: the imported stories with the document's
 * overrides applied. Without overrides this is the imported `sectionChrome`
 * as it is (null when the chrome is flat); with them, a flat chrome is first
 * spread over every section so the overridden one can differ from its
 * siblings. `parse` turns an override's JSON into a story (null skips it —
 * malformed attr data must not take anything down); callers cache there.
 */
export function effectiveSectionChrome<N>(
  doc: { attrs: Record<string, unknown> } | null,
  imported: ImportedChrome<N>,
  parse: (json: unknown) => N | null,
): ChromeStories<N>[] | null {
  if (!doc) return imported.sectionChrome;
  const overrides = doc.attrs['sectionChromeOverrides'] as Record<
    string,
    { headers?: Record<string, unknown>; footers?: Record<string, unknown> }
  > | null;
  if (!overrides || Object.keys(overrides).length === 0)
    return imported.sectionChrome;
  const count = Math.max(
    (doc.attrs['sections'] as unknown[] | null)?.length ?? 1,
    imported.sectionChrome?.length ?? 0,
  );
  const base: ChromeStories<N>[] =
    imported.sectionChrome ??
    Array.from({ length: count }, () => ({
      headers: imported.headers,
      footers: imported.footers,
      titlePg: imported.titlePg,
    }));
  const out = base.map((s) => ({
    headers: { ...s.headers },
    footers: { ...s.footers },
    titlePg: s.titlePg,
  }));
  for (const [key, o] of Object.entries(overrides)) {
    const target = out[Number(key)];
    if (!target) continue;
    for (const [variant, json] of Object.entries(o.headers ?? {})) {
      const parsed = parse(json);
      if (parsed) target.headers[variant] = parsed;
    }
    for (const [variant, json] of Object.entries(o.footers ?? {})) {
      const parsed = parse(json);
      if (parsed) target.footers[variant] = parsed;
    }
  }
  return out;
}
