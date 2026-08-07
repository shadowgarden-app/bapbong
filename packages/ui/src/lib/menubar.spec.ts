import { Collection } from '@shadow-garden/bapbong-contracts';
import type { Command } from '@shadow-garden/bapbong-contracts';
import { defaultMenus } from './menubar.js';

// Dropdown open/close + keyboard nav are verified in-browser (repo convention:
// package tests run in Node). Here we cover the pure default-menu structure.
const cmd = (name: string): Command => ({ name, run: () => false });

describe('defaultMenus', () => {
  it('builds a Format menu: marks, a separator, then alignments', () => {
    const commands = new Collection<Command>(
      [cmd('bold'), cmd('italic'), cmd('align-left'), cmd('align-center')],
      { idProperty: 'name' },
    );
    expect(defaultMenus(commands)).toEqual([
      {
        label: 'Format',
        entries: [
          { command: 'bold' },
          { command: 'italic' },
          'separator',
          { command: 'align-left' },
          { command: 'align-center' },
        ],
      },
    ]);
  });

  it('omits the separator when only one kind of command exists', () => {
    const commands = new Collection<Command>([cmd('bold'), cmd('underline')], {
      idProperty: 'name',
    });
    expect(defaultMenus(commands)).toEqual([
      {
        label: 'Format',
        entries: [{ command: 'bold' }, { command: 'underline' }],
      },
    ]);
  });
});
