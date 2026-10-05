import {
  Collection,
  KeybindingRegistry,
} from '@shadow-garden/bapbong-contracts';
import type { Command } from '@shadow-garden/bapbong-contracts';
import { defaultMenus, describeMenus } from './menubar.js';

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

describe('describeMenus', () => {
  it('lists every row by its menu path, with the label and shortcut the menubar shows', () => {
    const keys = new KeybindingRegistry(true);
    keys.add({ key: 'Mod-b', command: 'bold', source: 'core' });
    const app = new KeybindingRegistry(true);
    app.add({ key: 'Mod-s', command: 'save', source: 'app' });
    const lines = describeMenus(
      [
        {
          label: 'File',
          entries: [
            { label: 'Save', run: () => undefined, shortcutOf: 'save' },
            'separator',
            { label: 'Print', run: () => undefined, shortcut: '⌘P' },
          ],
        },
        {
          label: 'Format',
          entries: [
            {
              label: 'Text',
              submenu: [{ command: 'bold' }, { command: 'mystery' }],
            },
            { label: 'Margins', widget: () => null as unknown as HTMLElement },
          ],
        },
      ],
      { keybindings: [keys, app], labels: { mystery: 'Mystery' }, mac: true },
    );
    expect(lines).toEqual([
      'File ▸ Save  (⌘S)',
      'File ▸ Print  (⌘P)',
      'Format ▸ Text ▸ Bold  (⌘B)',
      'Format ▸ Text ▸ Mystery',
      'Format ▸ Margins',
    ]);
  });
});
