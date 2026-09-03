import { Collection } from '@shadow-garden/bapbong-contracts';
import type { Command } from '@shadow-garden/bapbong-contracts';
import {
  defaultToolbarGroups,
  foldPlan,
  type FoldCandidate,
} from './toolbar.js';

// DOM rendering + click behaviour is verified in-browser (repo convention: all
// package tests run in Node). Here we cover the pure grouping logic.
const cmd = (name: string): Command => ({ name, run: () => false });

describe('defaultToolbarGroups', () => {
  it('splits marks from alignments, preserving registry order', () => {
    const commands = new Collection<Command>(
      [cmd('bold'), cmd('italic'), cmd('align-left'), cmd('align-center')],
      { idProperty: 'name' },
    );
    expect(defaultToolbarGroups(commands)).toEqual([
      ['bold', 'italic'],
      ['align-left', 'align-center'],
    ]);
  });

  it('omits an empty group (no alignments registered)', () => {
    const commands = new Collection<Command>([cmd('bold'), cmd('underline')], {
      idProperty: 'name',
    });
    expect(defaultToolbarGroups(commands)).toEqual([['bold', 'underline']]);
  });
});

// foldPlan is the pure half of the overflow layout: one measurement batch in,
// the exact fold membership out. (DOM behaviour — anchor math, width caching,
// delta moves — is verified in-browser, per repo convention.)
describe('foldPlan', () => {
  const g = (
    idx: number,
    width: number,
    over: Partial<FoldCandidate> = {},
  ): FoldCandidate => ({ idx, width, sticky: false, hidden: false, ...over });

  it('folds nothing while the row fits, gaps included', () => {
    // 3×100 + 2×10 gap = 320 exactly.
    expect(foldPlan([g(0, 100), g(1, 100), g(2, 100)], 320, 24, 10).size).toBe(
      0,
    );
    // One px less and the tail folds — and must now clear room for ⋮ too.
    expect(foldPlan([g(0, 100), g(1, 100), g(2, 100)], 318, 24, 10)).toEqual(
      new Set([2]),
    );
  });

  it('folds tail-first among ordinary groups', () => {
    const plan = foldPlan(
      [g(0, 100), g(1, 100), g(2, 100), g(3, 100)],
      250,
      24,
      10,
    );
    expect(plan).toEqual(new Set([2, 3]));
  });

  it('a sticky group outlasts every ordinary one, wherever it sits', () => {
    const cands = [
      g(0, 100),
      g(1, 100, { sticky: true }),
      g(2, 100),
      g(3, 100),
    ];
    expect(foldPlan(cands, 250, 24, 10)).toEqual(new Set([3, 2]));
    // Tighter: ordinary groups 3, 2, then 0 fold before the sticky 1 does.
    expect(foldPlan(cands, 150, 24, 10)).toEqual(new Set([3, 2, 0]));
    // Only when nothing else can give way does sticky fold too.
    expect(foldPlan(cands, 60, 24, 10)).toEqual(new Set([3, 2, 0, 1]));
  });

  it('sticky groups fold tail-first among themselves', () => {
    const cands = [
      g(0, 100, { sticky: true }),
      g(1, 100),
      g(2, 100, { sticky: true }),
    ];
    expect(foldPlan(cands, 150, 24, 10)).toEqual(new Set([1, 2]));
  });

  it('hidden groups take no space and never fold', () => {
    const cands = [g(0, 100), g(1, 400, { hidden: true }), g(2, 100)];
    expect(foldPlan(cands, 210, 24, 10).size).toBe(0);
    expect(foldPlan(cands, 150, 24, 10)).toEqual(new Set([2]));
  });
});
