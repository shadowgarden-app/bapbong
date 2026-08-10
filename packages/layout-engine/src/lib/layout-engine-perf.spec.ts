import { perf } from '@shadow-garden/bapbong-contracts';
import type {
  FlowBlock,
  FlowParagraph,
  FlowTableCell,
  FontSpec,
  LayoutConfig,
  MeasureText,
} from '@shadow-garden/bapbong-contracts';
import { layoutBlocks } from './layout-engine.js';

/**
 * Pagination cost guards for long tables.
 *
 * A table spanning K pages used to deep-clone every row below the cut once
 * per page — O(rows × K) — in splitTableAt's remainder, and again in every
 * per-page checkpoint. The TableRemainder cursor makes fragment construction
 * clone each row once: these tests pin that down so the quadratic can't come
 * back silently.
 */

let measureCalls = 0;
const measure: MeasureText = (text) => {
  measureCalls++;
  return text.length * 10;
};
const font = (): FontSpec => ({
  family: 'Arial',
  sizePt: 10,
  bold: false,
  italic: false,
});
const para = (text: string): FlowParagraph => ({
  type: 'paragraph',
  runs: [{ text, font: font() }],
});
const config = (h: number): LayoutConfig => ({
  measureText: measure,
  defaultFont: { sizePt: 10 },
  page: {
    width: 400,
    height: h,
    margin: { top: 20, right: 20, bottom: 20, left: 20 },
  },
});
const cellOf = (content: FlowBlock[]): FlowTableCell => ({
  colspan: 1,
  rowspan: 1,
  colwidth: null,
  content,
});

const ROWS = 400;
const COLS = 3;
const bigTable = (): FlowBlock => ({
  type: 'table',
  rows: Array.from({ length: ROWS }, (_, r) => ({
    cells: Array.from({ length: COLS }, (_, c) =>
      cellOf([para(`row${r} cell${c} aaa bbb ccc ddd`)]),
    ),
  })),
});

describe('table pagination cost', () => {
  it('splitting a long table stays measurement-free', () => {
    // All measureText calls belong to the wrap; pagination must re-measure
    // nothing no matter how many page boundaries cut the table.
    measureCalls = 0;
    layoutBlocks([bigTable()], config(5000)); // few pages, ~no splits
    const tall = measureCalls;
    measureCalls = 0;
    layoutBlocks([bigTable()], config(160)); // a split at every page
    const short = measureCalls;
    expect(short).toBe(tall);
  });

  it('fragment construction clones each row once, not once per page', () => {
    // remainderView reports every cell it clones. Across the whole table the
    // total must stay linear in the cell count: each cell lands in exactly one
    // fragment (straddlers add a handful at the page seams). The old
    // representation cloned every below-cut row once per page — for this
    // geometry ~60× the linear bound — so a reintroduced quadratic trips the
    // assertion loudly, with head-room for splitter bookkeeping.
    perf.setEnabled(true);
    try {
      const g = globalThis as unknown as {
        __BAPBONG_PERF_STATE__?: { counters: Record<string, number> };
      };
      const counters = () => g.__BAPBONG_PERF_STATE__?.counters ?? {};
      delete counters()['table.view.cellsCloned'];
      const { pages } = layoutBlocks([bigTable()], config(160));
      expect(pages.length).toBeGreaterThan(300); // sanity: real splitting ran
      const cloned = counters()['table.view.cellsCloned'] ?? 0;
      expect(cloned).toBeGreaterThan(0);
      expect(cloned).toBeLessThanOrEqual(ROWS * COLS * 3);
    } finally {
      perf.setEnabled(false);
    }
  });

  it('heavy splitting costs the same order as laying out, not multiples', () => {
    // Timing bound with deliberate slack for CI jitter: the split-per-page
    // layout may cost a few times the no-split layout (fragments still get
    // built), but the old quadratic put it at ~8× and growing with row count.
    const t0 = performance.now();
    layoutBlocks([bigTable()], config(5000));
    const tallMs = performance.now() - t0;
    const t1 = performance.now();
    layoutBlocks([bigTable()], config(160));
    const shortMs = performance.now() - t1;
    expect(shortMs).toBeLessThan(Math.max(60, tallMs * 4));
  });
});
