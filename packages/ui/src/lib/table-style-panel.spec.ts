import { describe, expect, it } from 'vitest';
import { tblLookBits } from './table-style-panel.js';

// DOM rendering is verified in-browser (repo convention: package tests run in
// Node). The pure piece here is the bitmask the footer shows — it must match
// what the exporter writes for the same look.
describe('tblLookBits', () => {
  it('spells Word’s bitmask (noHBand/noVBand are negative bits)', () => {
    expect(
      tblLookBits({
        firstRow: true,
        lastRow: false,
        firstCol: true,
        lastCol: false,
        hBand: true,
        vBand: false,
      }),
    ).toBe('04A0');
    expect(
      tblLookBits({
        firstRow: true,
        lastRow: true,
        firstCol: true,
        lastCol: true,
        hBand: true,
        vBand: true,
      }),
    ).toBe('01E0');
    expect(
      tblLookBits({
        firstRow: false,
        lastRow: false,
        firstCol: false,
        lastCol: false,
        hBand: false,
        vBand: false,
      }),
    ).toBe('0600');
  });
});
