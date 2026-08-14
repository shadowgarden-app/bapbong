import fs from 'node:fs';
import path from 'node:path';
import {
  importDocx,
  layout,
  FontRegistry,
  createFontRegistryMeasurer,
  createFontRegistryMetrics,
  createApproxMeasurer,
  createApproxMetrics,
} from './index.js';

/**
 * Pagination measured against WORD, not against ourselves.
 *
 * Every page-count baseline in this repo used to be self-referential: a change
 * was "safe" when our own numbers did not move, which says nothing about
 * whether they were right. These are the counts Microsoft Word reports for the
 * same files (read off its status bar, 2026-08-14), so a layout change can
 * finally be judged by whether it moves TOWARD them.
 *
 * `ours` is what this engine produces today. Where it differs from `word`, the
 * gap is a known defect with a direction, not a mystery:
 *
 *   - large_sample, de_cuong — exact. Both are ordinary flowed text, which is
 *     the case the engine has been tuned against.
 *   - khbd +5. It was 31 SHORT until the font fallback landed: the document is
 *     96% Calibri Light, a family no registry carries, and everything about it
 *     — widths and line heights — was coming from the approximate measurer.
 *     It now borrows Calibri's vertical metrics (not its widths).
 *   - LeMinhThu −3, bc_rieng −2 — both Times New Roman, so neither is a font
 *     problem; both are table-heavy, and that is where to look next.
 *   - NAVY +2 — the one multi-column, float-heavy document, and the one where
 *     we run LONG.
 *
 * When a fix lands, update `ours` in the same commit and say which way it
 * moved. Do not update `word`: those are the target.
 */
const BASELINE: { file: string; word: number; ours: number }[] = [
  { file: 'large_sample.docx', word: 122, ours: 122 },
  { file: 'de_cuong_cuoi_ki.docx', word: 7, ours: 7 },
  { file: 'bc_rieng.docx', word: 28, ours: 26 },
  { file: 'LeMinhThu52DL- Shop Kpop.docx', word: 22, ours: 19 },
  { file: 'khbd.docx', word: 246, ours: 251 },
  {
    file: 'NAVY HOTEL THE REGION IV- Factsheet-Vietnamese.docx',
    word: 3,
    ours: 5,
  },
];

const PUB = path.resolve(__dirname, '../../../apps/playground/public');
const FONTS = path.resolve(__dirname, '../../../node_modules/@fontsource');

/** The bundled metric-compatible faces the app itself measures with — without
 *  them these counts would describe the test runner's fonts, not Word's. */
function registry(): FontRegistry {
  const r = new FontRegistry();
  for (const { file, names } of [
    { file: 'carlito', names: ['Carlito', 'Calibri'] },
    { file: 'tinos', names: ['Tinos', 'Times New Roman'] },
    { file: 'arimo', names: ['Arimo', 'Arial'] },
  ])
    for (const w of ['400', '700'])
      for (const st of ['normal', 'italic'])
        for (const sub of ['latin', 'latin-ext', 'vietnamese']) {
          const p = `${FONTS}/${file}/files/${file}-${sub}-${w}-${st}.woff`;
          if (!fs.existsSync(p)) continue;
          const b = fs.readFileSync(p);
          for (const n of names)
            r.registerBytes(
              n,
              { bold: w === '700', italic: st === 'italic' },
              b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength),
            );
        }
  return r;
}

describe('pagination against Word', () => {
  const reg = registry();
  const cfg = {
    measureText: createFontRegistryMeasurer(reg, createApproxMeasurer()),
    measureMetrics: createFontRegistryMetrics(reg, createApproxMetrics()),
  };

  for (const { file, word, ours } of BASELINE) {
    it(`${file} — Word ${word}, ours ${ours}`, async () => {
      const bytes = fs.readFileSync(path.join(PUB, file));
      const { doc, page } = await importDocx(
        bytes.buffer.slice(
          bytes.byteOffset,
          bytes.byteOffset + bytes.byteLength,
        ) as ArrayBuffer,
      );
      const pages = layout(doc, { page, ...cfg }).pages.length;
      // The assertion is on `ours` — the test's job is to make a change in
      // pagination impossible to land silently. `word` rides along so the
      // direction of any change is visible in the same line.
      expect(pages).toBe(ours);
    });
  }
});
