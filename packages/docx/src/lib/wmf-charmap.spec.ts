import { decodeWmfText } from './wmf-charmap.js';

const bytes = (...b: number[]): Uint8Array => new Uint8Array(b);

describe('decodeWmfText', () => {
  it('maps Symbol bytes: greek, operators, and mixed ASCII passthrough', () => {
    // "w = 2pf" the way MathType draws ω = 2πf in Symbol runs.
    expect(decodeWmfText(bytes(0x77, 0x3d, 0x70), 'Symbol')).toBe('ω=π');
    expect(decodeWmfText(bytes(0x44, 0xce, 0xa5, 0xb9), 'Symbol')).toBe('Δ∈∞≠');
  });

  it('maps Symbol delimiter pieces to the Unicode piece characters', () => {
    // Tall parens are assembled top/extender/bottom.
    expect(decodeWmfText(bytes(0xe6, 0xe7, 0xe8), 'Symbol')).toBe('⎛⎜⎝');
    expect(decodeWmfText(bytes(0xf6, 0xf7, 0xf8), 'Symbol')).toBe('⎞⎟⎠');
  });

  it('maps MT Extra vector embellishments and set letters', () => {
    // The arrow drawn over B in a vector: extender piece + combining head.
    expect(decodeWmfText(bytes(0x75, 0x72), 'MT Extra')).toBe('⎯⃗');
    // Double-struck ℝ/ℤ, verified against the files' embedded MTEF.
    expect(decodeWmfText(bytes(0xa1), 'MT Extra')).toBe('ℝ');
    expect(decodeWmfText(bytes(0xa2), 'MT Extra')).toBe('ℤ');
  });

  it('decodes other faces as windows-1252, not latin-1', () => {
    expect(
      decodeWmfText(bytes(0x66, 0x28, 0x78, 0x29), 'Times New Roman'),
    ).toBe('f(x)');
    // 0x92 is a curly apostrophe in cp1252, a C1 control in latin-1.
    expect(decodeWmfText(bytes(0x92), 'Times New Roman')).toBe('’');
    expect(decodeWmfText(bytes(0xd0), 'Times New Roman')).toBe('Ð');
  });

  it('turns unmapped symbol-font bytes into U+FFFD, never a wrong glyph', () => {
    expect(decodeWmfText(bytes(0xf0), 'Symbol')).toBe('�');
    expect(decodeWmfText(bytes(0x21), 'MT Extra')).toBe('�');
  });
});
