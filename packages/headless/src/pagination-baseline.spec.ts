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
 *   - de_cuong, bc_rieng — exact. bc_rieng was −2 until empty paragraphs
 *     started measuring their own paragraph mark instead of the document
 *     default font.
 *   - LeMinhThu −1, from −3 for the same reason. Both it and bc_rieng are
 *     table-heavy Times New Roman documents, and both moved the moment the
 *     mark was measured.
 *   - large_sample +1 and khbd +11 moved the OTHER way on that same change,
 *     and the change was not the error: khbd's blank lines really are 13pt
 *     Calibri Light (1230 of them say so on the mark), we really did draw them
 *     11pt, and 1145 of them sit inside table cells. Drawing them too short
 *     had been hiding something that makes our pages hold less than Word's —
 *     almost certainly the table row geometry khbd is already known for.
 *     Neither number is evidence against the mark; both point at rows.
 *   - NAVY +2 — the one multi-column, float-heavy document. Two of its five
 *     pages are near-blank, and they come from section breaks, not from
 *     measurement: a top-margin-only change at a continuous break splits one
 *     page in two here (removing it alone gives 4), and the last empty
 *     paragraph of a section spills off page 1.
 *   - nested_table −2, and the one entry these numbers cannot yet judge: it is
 *     75% Open Sans and 24% Aptos. We bundle neither, this Mac has neither
 *     installed, and Aptos — Microsoft 365's default since 2023 — has no
 *     metric-compatible clone in existence. Headless it measures through the
 *     approximation end to end (3 pages); in the browser, where a substitute
 *     face at least supplies real ascent and descent, it reaches 4.
 *
 * When a fix lands, update `ours` in the same commit and say which way it
 * moved. Do not update `word`: those are the target.
 */
const BASELINE: { file: string; word: number; ours: number }[] = [
  { file: 'large_sample.docx', word: 122, ours: 123 },
  { file: 'de_cuong_cuoi_ki.docx', word: 7, ours: 7 },
  { file: 'bc_rieng.docx', word: 28, ours: 28 },
  { file: 'LeMinhThu52DL- Shop Kpop.docx', word: 22, ours: 21 },
  { file: 'khbd.docx', word: 246, ours: 257 },
  {
    file: 'NAVY HOTEL THE REGION IV- Factsheet-Vietnamese.docx',
    word: 3,
    ours: 5,
  },
  { file: 'nested_table.docx', word: 5, ours: 3 },
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
