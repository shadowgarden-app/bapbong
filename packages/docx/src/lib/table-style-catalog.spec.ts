import { describe, expect, it } from 'vitest';
import { cellStyleLayer } from '@shadow-garden/bapbong-contracts';
import {
  catalogStyleXml,
  catalogTableStyles,
  TABLE_STYLE_CATALOG,
} from './table-style-catalog';
import { parseXml } from './ooxml';

describe('the built-in table-style catalog', () => {
  it('is thirteen well-formed Word-named styles', () => {
    expect(TABLE_STYLE_CATALOG).toHaveLength(13);
    for (const e of TABLE_STYLE_CATALOG) {
      // parseXml throws on malformed input; the id must appear verbatim so
      // the exporter's "does styles.xml already define it" check works.
      parseXml(`<w:styles xmlns:w="x">${e.xml}</w:styles>`);
      expect(e.xml).toContain(`w:styleId="${e.id}"`);
      expect(e.xml).toContain(`w:name w:val="${e.name}"`);
    }
    expect(catalogStyleXml('LightGrid-Accent3')).toContain('9BBB59');
    expect(catalogStyleXml('Nope')).toBeUndefined();
  });

  it('resolves through the import pipeline into working styles', () => {
    const styles = catalogTableStyles();
    expect(styles).toHaveLength(13);
    const lg = styles.find((s) => s.id === 'LightGrid-Accent3')?.style;
    if (!lg) throw new Error('LightGrid-Accent3 missing');
    // Word's own accent-3 band literal is D6E3BC (the Kpop corpus file);
    // its tint math is not a linear mix, so one factor cannot land all
    // three channels exactly. Within one step per channel is invisible —
    // and whatever we emit here is also what our XML makes Word render.
    const band = lg.cond.band1Horz?.background ?? '#000000';
    const target = [0xd6, 0xe3, 0xbc];
    [1, 3, 5].forEach((i, ch) =>
      expect(
        Math.abs(parseInt(band.slice(i, i + 2), 16) - target[ch]),
      ).toBeLessThanOrEqual(1),
    );
    expect(lg.cond.firstRow?.font?.bold).toBe(true);
    expect(lg.bands).toEqual({ row: 1, col: 1 });
    expect(lg.table.cellPadding).toEqual({
      top: 0,
      left: 7,
      bottom: 0,
      right: 7,
    });
    // Medium Shading: solid accent header with white bold text.
    const ms = styles.find((s) => s.id === 'MediumShading1-Accent1')?.style;
    expect(ms?.cond.firstRow?.background).toBe('#4F81BD');
    expect(ms?.cond.firstRow?.font).toMatchObject({
      bold: true,
      color: '#FFFFFF',
    });
    // Table Grid styles nothing conditionally.
    const tg = styles.find((s) => s.id === 'TableGrid')?.style;
    expect(Object.keys(tg?.cond ?? { x: 1 })).toHaveLength(0);
  });

  it('a header cell comes out styled through cellStyleLayer', () => {
    const lg = catalogTableStyles().find(
      (s) => s.id === 'LightGrid-Accent1',
    )?.style;
    if (!lg) throw new Error('missing');
    const layer = cellStyleLayer(
      lg,
      {
        firstRow: true,
        lastRow: false,
        firstCol: true,
        lastCol: false,
        hBand: true,
        vBand: false,
      },
      { row: 0, rowCount: 4, col: 1, colspan: 1, colCount: 4, header: false },
    );
    expect(layer.font?.bold).toBe(true);
    const bottom = layer.borders?.bottom;
    if (!bottom) throw new Error('header bottom edge missing');
    expect(bottom.width).toBeGreaterThan(1);
    // The SHEET keeps the branch's explicit insideH=nil (region semantics);
    // a one-row region has no interior horizontal edge, so the per-cell
    // layer only ever emits the four physical sides.
    expect(lg.cond.firstRow?.borders?.insideH).toBe(false);
    expect(Object.keys(layer.borders ?? {}).sort()).toEqual([
      'bottom',
      'left',
      'right',
      'top',
    ]);
  });
});
