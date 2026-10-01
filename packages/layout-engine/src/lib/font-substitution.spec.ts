import { schema } from '@shadow-garden/bapbong-model';
import type {
  FontAvailable,
  FontSubstitutes,
  LayoutConfig,
  LayoutLine,
  MeasureText,
} from '@shadow-garden/bapbong-contracts';
import { fontFamilyResolver, layout } from './layout-engine.js';

const SUBS: FontSubstitutes = {
  'Arial MT': { altNames: ['Arial'], generic: 'swiss' },
  'Zz Sans': { generic: 'swiss' },
  'Zz Serif': { altNames: ['Nowhere Serif'], generic: 'roman' },
  'Zz Mono': { generic: 'roman', fixedPitch: true },
  'VNI-Times': { altNames: ['Times New Roman'] },
};

/** This host draws only the families listed. */
const only =
  (...families: string[]): FontAvailable =>
  (family) =>
    families.includes(family);

describe('fontFamilyResolver', () => {
  it('draws a missing font in the first available altName', () => {
    const drawn = fontFamilyResolver(SUBS, only('Arial'));
    expect(drawn('Arial MT')).toBe('Arial');
    // Word matches font names case-insensitively.
    expect(drawn('arial mt')).toBe('Arial');
  });

  it('falls back to the face of the font’s class, fixed pitch first', () => {
    const drawn = fontFamilyResolver(
      SUBS,
      only('Arial', 'Times New Roman', 'Courier New'),
    );
    expect(drawn('Zz Sans')).toBe('Arial');
    expect(drawn('Zz Serif')).toBe('Times New Roman');
    expect(drawn('Zz Mono')).toBe('Courier New');
  });

  it('keeps a font that is available, whatever its table says', () => {
    const drawn = fontFamilyResolver(SUBS, only('VNI-Times', 'Arial MT'));
    expect(drawn('VNI-Times')).toBe('VNI-Times');
    expect(drawn('Arial MT')).toBe('Arial MT');
  });

  it('keeps the name when no substitute is available either', () => {
    expect(fontFamilyResolver(SUBS, only())('Arial MT')).toBe('Arial MT');
  });

  it('keeps fonts the table does not describe', () => {
    expect(fontFamilyResolver(SUBS, only('Arial'))('Foo')).toBe('Foo');
  });

  it('substitutes nothing without a way to ask what is available', () => {
    expect(fontFamilyResolver(SUBS, undefined)('Arial MT')).toBe('Arial MT');
    expect(fontFamilyResolver(null, only('Arial'))('Arial MT')).toBe(
      'Arial MT',
    );
  });
});

describe('layout with font substitutes', () => {
  // Arial measures 5px a character, anything else 10px — so a width tells
  // which family the layout measured with.
  const measure: MeasureText = (text, font) =>
    text.length * (font.family === 'Arial' ? 5 : 10);

  const config = (fontAvailable?: FontAvailable): LayoutConfig => ({
    measureText: measure,
    page: {
      width: 400,
      height: 400,
      margin: { top: 20, right: 20, bottom: 20, left: 20 },
    },
    ...(fontAvailable && { fontAvailable }),
  });

  const doc = (subs: FontSubstitutes | null) =>
    schema.node('doc', { fontSubstitutes: subs }, [
      schema.node('paragraph', null, [
        schema.text('3500', [
          schema.marks['fontFamily'].create({ family: 'Arial MT' }),
        ]),
      ]),
    ]);

  const firstSegment = (lines: LayoutLine[]) => lines[0].segments[0];

  it('measures and paints a missing font in its substitute', () => {
    const { pages } = layout(doc(SUBS), config(only('Arial')));
    const seg = firstSegment(pages[0].lines);
    expect(seg.font.family).toBe('Arial');
    expect(seg.width).toBe(20);
  });

  it('keeps the named font where it is available', () => {
    const { pages } = layout(doc(SUBS), config(only('Arial', 'Arial MT')));
    const seg = firstSegment(pages[0].lines);
    expect(seg.font.family).toBe('Arial MT');
    expect(seg.width).toBe(40);
  });

  it('is unchanged for a document without a font table', () => {
    const { pages } = layout(doc(null), config(only('Arial')));
    expect(firstSegment(pages[0].lines).font.family).toBe('Arial MT');
  });

  it('leaves the model untouched', () => {
    const d = doc(SUBS);
    layout(d, config(only('Arial')));
    expect(d.firstChild!.firstChild!.marks[0].attrs['family']).toBe('Arial MT');
  });
});
