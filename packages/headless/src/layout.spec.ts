import * as opentype from 'opentype.js';
// Everything comes through the façade — this is the surface a backend installs.
// A missing re-export makes this import (and the test) fail loudly.
import {
  schema,
  layout,
  FontRegistry,
  createFontRegistryMeasurer,
  createFontRegistryMetrics,
  createApproxMeasurer,
  createApproxMetrics,
  type PageConfig,
} from './index.js';

/** A synthetic font (parsed, so opentype's cmap works) covering A–Z + space with
 *  a fixed advance, so pagination is exactly controlled and platform-independent. */
function fontBytes(advance: number): ArrayBuffer {
  const glyphs = [new opentype.Glyph({ name: '.notdef', unicode: 0, advanceWidth: advance, path: new opentype.Path() })];
  for (let c = 65; c <= 90; c++) glyphs.push(new opentype.Glyph({ name: `u${c}`, unicode: c, advanceWidth: advance, path: new opentype.Path() }));
  glyphs.push(new opentype.Glyph({ name: 'space', unicode: 32, advanceWidth: advance, path: new opentype.Path() }));
  return new opentype.Font({ familyName: 'TestFont', styleName: 'Regular', unitsPerEm: 1000, ascender: 800, descender: -200, glyphs }).toArrayBuffer();
}

function registry(advance: number): FontRegistry {
  const reg = new FontRegistry();
  reg.registerBytes('TestFont', {}, fontBytes(advance));
  return reg;
}

const PAGE: PageConfig = { width: 794, height: 1123, margin: { top: 96, right: 96, bottom: 96, left: 96 } };

/** A doc of `n` identical paragraphs, all rendered in the (default) TestFont. */
function longDoc(n: number) {
  const line = 'THE QUICK BROWN FOX JUMPS OVER THE LAZY DOG';
  const paras = Array.from({ length: n }, () => schema.node('paragraph', null, [schema.text(line)]));
  return schema.node('doc', null, paras);
}

const cfg = (reg: FontRegistry) => ({
  page: PAGE,
  measureText: createFontRegistryMeasurer(reg, createApproxMeasurer()),
  measureMetrics: createFontRegistryMetrics(reg, createApproxMetrics()),
  defaultFont: { family: 'TestFont', sizePt: 12 },
});

describe('headless pagination (DOM-free, font-metric)', () => {
  it('paginates a document in Node using font-file metrics', () => {
    const resolved = layout(longDoc(120), cfg(registry(500)));
    expect(resolved.pages.length).toBeGreaterThan(1);
  });

  it('is deterministic — identical inputs give identical page counts', () => {
    const reg = registry(500);
    const a = layout(longDoc(120), cfg(reg));
    const b = layout(longDoc(120), cfg(reg));
    expect(b.pages.length).toBe(a.pages.length);
  });

  it('font metrics drive layout — wider advances wrap sooner → more pages', () => {
    const narrow = layout(longDoc(120), cfg(registry(400)));
    const wide = layout(longDoc(120), cfg(registry(1000)));
    expect(wide.pages.length).toBeGreaterThan(narrow.pages.length);
  });
});
