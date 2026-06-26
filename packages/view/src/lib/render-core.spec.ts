import { schema } from '@shadow-garden/bapbong-model';
import { RenderCore, BapbongView, collectFontFamilies } from '../index.js';

// RenderCore / BapbongView wire a canvas + DOM, exercised end-to-end in the
// playground; here we assert the public surface plus the pure helpers (the
// node test env has no DOM, so construction is covered in-browser).
describe('bapbong-view exports', () => {
  it('exposes RenderCore + BapbongView as classes', () => {
    expect(typeof RenderCore).toBe('function');
    expect(typeof BapbongView).toBe('function');
  });
});

describe('collectFontFamilies', () => {
  it('always includes the engine default (Arial)', () => {
    const doc = schema.node('doc', null, [schema.node('paragraph', null, [schema.text('hi')])]);
    expect(collectFontFamilies(doc)).toContain('Arial');
  });

  it('gathers distinct fontFamily marks across documents', () => {
    const para = (text: string, family: string) =>
      schema.node('paragraph', null, [schema.text(text, [schema.marks['fontFamily'].create({ family })])]);
    const a = schema.node('doc', null, [para('x', 'Calibri')]);
    const b = schema.node('doc', null, [para('y', 'Times New Roman'), para('z', 'Calibri')]);
    const fams = collectFontFamilies(a, b, undefined);
    expect(fams).toContain('Calibri');
    expect(fams).toContain('Times New Roman');
    expect(fams).toContain('Arial');
    // distinct (Calibri appears in both docs but once in the result)
    expect(fams.filter((f) => f === 'Calibri')).toHaveLength(1);
  });
});
