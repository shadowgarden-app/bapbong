import { Collection } from '@shadow-garden/bapbong-contracts';
import type { Command } from '@shadow-garden/bapbong-contracts';
import { defaultToolbarGroups } from './toolbar.js';

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
