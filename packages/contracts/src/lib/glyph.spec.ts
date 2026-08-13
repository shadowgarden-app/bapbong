import type { FontSpec } from './contracts.js';
import {
  applyGlyphSpec,
  glyphCount,
  glyphKey,
  sameGlyphRun,
  type GlyphContext,
} from './glyph.js';

const font = (over: Partial<FontSpec> = {}): FontSpec => ({
  family: 'Arial',
  sizePt: 11,
  bold: false,
  italic: false,
  ...over,
});

const fakeCtx = (): GlyphContext => ({
  font: '',
  letterSpacing: '',
  fontKerning: '',
});

describe('applyGlyphSpec', () => {
  it('writes the shorthand and the tracking', () => {
    const ctx = fakeCtx();
    applyGlyphSpec(ctx, font({ bold: true, letterSpacing: 1.3 }));
    expect(ctx.font).toBe('700 11pt Arial');
    expect(ctx.letterSpacing).toBe('1.3px');
  });

  it('always writes letterSpacing, so a context is never left tracked', () => {
    // Contexts are reused across runs; leaving the previous run's tracking in
    // place would silently widen the next one.
    const ctx = fakeCtx();
    applyGlyphSpec(ctx, font({ letterSpacing: 4 }));
    applyGlyphSpec(ctx, font());
    expect(ctx.letterSpacing).toBe('0px');
  });

  it('kerns only when a run opted in, and always says which', () => {
    // Word's default is NO kerning ("if this element is never applied in the
    // style hierarchy, then font kerning shall not be applied"), so a spec
    // that says nothing must come out 'none'. Written both ways round for the
    // reuse hazard: a context left at 'normal' from a previous run would
    // silently narrow the next one.
    const ctx = fakeCtx();
    applyGlyphSpec(ctx, font());
    expect(ctx.fontKerning).toBe('none');
    applyGlyphSpec(ctx, font({ kerning: true }));
    expect(ctx.fontKerning).toBe('normal');
    applyGlyphSpec(ctx, font({ kerning: false }));
    expect(ctx.fontKerning).toBe('none');
  });

  it('configures measurer and painter identically — the whole point', () => {
    // The measurer and the painter live in packages that may not import each
    // other; they agree only because both call this. Same spec in, same
    // context state out, so the browser does the arithmetic once for both.
    const measurerCtx = fakeCtx();
    const painterCtx = fakeCtx();
    const spec = font({ italic: true, sizePt: 9, letterSpacing: 2.5 });
    applyGlyphSpec(measurerCtx, spec);
    applyGlyphSpec(painterCtx, spec);
    expect(painterCtx).toEqual(measurerCtx);
  });
});

describe('applyGlyphSpec + horizontal scale', () => {
  it('pre-divides tracking so measurer and painter both land on it', () => {
    // The two callers apply the scale differently — the measurer multiplies
    // the width it measured, the painter puts it on the transform — and this
    // pre-division is what makes both come out at `glyphs × scale + tracking`.
    const ctx = fakeCtx();
    const scale = applyGlyphSpec(ctx, font({ scaleX: 0.8, letterSpacing: 2 }));
    expect(scale).toBe(0.8);
    expect(ctx.letterSpacing).toBe('2.5px'); // 2 / 0.8

    // What the measurer computes: (glyphs + n × 2.5) × 0.8.
    const glyphs = 100;
    const n = 4;
    const measured = (glyphs + n * 2.5) * 0.8;
    // Word squeezes the glyphs and leaves the tracking alone.
    expect(measured).toBeCloseTo(glyphs * 0.8 + n * 2, 10);
  });

  it('returns 1 and leaves tracking alone when unscaled', () => {
    const ctx = fakeCtx();
    expect(applyGlyphSpec(ctx, font({ letterSpacing: 2 }))).toBe(1);
    expect(ctx.letterSpacing).toBe('2px');
  });
});

describe('sameGlyphRun', () => {
  it('separates runs that differ only in tracking', () => {
    // The cluster merge in the layout engine measures consecutive same-run
    // tokens cumulatively. Treating these two as one run would measure the
    // joined text with one side's tracking.
    expect(sameGlyphRun(font({ letterSpacing: 2 }), font())).toBe(false);
    expect(
      sameGlyphRun(font({ letterSpacing: 2 }), font({ letterSpacing: 2 })),
    ).toBe(true);
  });

  it('separates runs that differ only in horizontal scale', () => {
    expect(sameGlyphRun(font({ scaleX: 0.8 }), font())).toBe(false);
    expect(sameGlyphRun(font({ scaleX: 0.8 }), font({ scaleX: 0.8 }))).toBe(
      true,
    );
  });

  it('still separates the face fields', () => {
    expect(sameGlyphRun(font(), font({ bold: true }))).toBe(false);
    expect(sameGlyphRun(font(), font({ sizePt: 12 }))).toBe(false);
    expect(sameGlyphRun(font(), font({ family: 'Times' }))).toBe(false);
    expect(sameGlyphRun(font(), font({ italic: true }))).toBe(false);
  });

  it('agrees with glyphKey', () => {
    const a = font({ letterSpacing: 2 });
    const b = font({ letterSpacing: 2 });
    expect(glyphKey(a)).toBe(glyphKey(b));
    expect(glyphKey(a)).not.toBe(glyphKey(font()));
  });
});

describe('glyphCount', () => {
  it('counts code points, not UTF-16 units', () => {
    expect(glyphCount('abc')).toBe(3);
    expect(glyphCount('a😀b')).toBe(3); // surrogate pair is one character
  });

  it('counts precomposed Vietnamese as one unit per letter', () => {
    // The corpus is NFC, where this matches what a browser charges tracking
    // for. Built with normalize() rather than pasted literals: the two forms
    // of a word look identical in source, so the intent has to live in the
    // code. Decomposed sequences count higher — documented drift, and only
    // the arithmetic measurers use this at all.
    const word = 'Ti\u1ebfng'.normalize('NFC');
    expect(glyphCount(word)).toBe(5);
    expect(glyphCount(word.normalize('NFD'))).toBeGreaterThan(5);
  });
});
