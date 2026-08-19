import {
  KeybindingRegistry,
  filterKeybindingRows,
  formatKey,
  hasCommandModifier,
  isValidKey,
  keyLabel,
  keybindingRows,
  normalizeKey,
  type Keybinding,
} from './keybinding.js';

describe('keybinding grammar', () => {
  it("normalizes to ProseMirror's event shape, Mod per platform", () => {
    expect(normalizeKey('Mod-Shift-z', true)).toBe('Meta-Shift-z');
    expect(normalizeKey('Shift-Mod-z', true)).toBe('Meta-Shift-z'); // order-free
    expect(normalizeKey('Mod-Shift-z', false)).toBe('Ctrl-Shift-z');
    expect(normalizeKey('Mod-B', true)).toBe('Meta-b'); // letters lower-case
    expect(normalizeKey('Shift-Mod-Z', true)).toBe('Meta-Shift-z');
    expect(normalizeKey('Alt-x', false)).toBe('Alt-x');
    expect(normalizeKey('Enter', true)).toBe('Enter');
    expect(normalizeKey('Shift-Tab', true)).toBe('Shift-Tab');
    expect(normalizeKey('Mod-\\', true)).toBe('Meta-\\');
    expect(() => normalizeKey('Cmdd-b', true)).toThrow(RangeError);
  });

  it('knows which chords may live at window scope', () => {
    expect(hasCommandModifier('Mod-s')).toBe(true);
    expect(hasCommandModifier('Alt-x')).toBe(true);
    expect(hasCommandModifier('Shift-Enter')).toBe(false);
    expect(hasCommandModifier('Enter')).toBe(false);
    expect(isValidKey('Mod-Shift-z')).toBe(true);
    expect(isValidKey('Mod+z')).toBe(false); // wrong separator
    expect(isValidKey('Foo-z')).toBe(false);
  });

  it('formats for the platform, in its order and glyphs', () => {
    expect(formatKey('Mod-Shift-z', true)).toEqual(['⇧', '⌘', 'Z']);
    expect(formatKey('Ctrl-Alt-Shift-Mod-b', true)).toEqual([
      '⌃',
      '⌥',
      '⇧',
      '⌘',
      'B',
    ]);
    expect(formatKey('Mod-Shift-z', false)).toEqual(['Ctrl', 'Shift', 'Z']);
    expect(formatKey('Alt-x', false)).toEqual(['Alt', 'X']);
    expect(formatKey('Enter', true)).toEqual(['↩']);
    expect(formatKey('Shift-Tab', false)).toEqual(['Shift', 'Tab']);
    expect(keyLabel('Mod-Shift-z', true)).toBe('⇧⌘Z');
    expect(keyLabel('Mod-Shift-z', false)).toBe('Ctrl+Shift+Z');
  });

  it('builds sorted rows per command and filters across every column', () => {
    const bindings: Keybinding[] = [
      { key: 'Mod-z', command: 'undo', source: 'core' },
      { key: 'Shift-Mod-z', command: 'redo', source: 'core' },
      { key: 'Mod-y', command: 'redo', source: 'core' },
      { key: 'Tab', command: 'list-indent', when: 'in a list', source: 'core' },
      { key: 'Mod-s', command: 'save', scope: 'window', source: 'desktop' },
    ];
    const titles: Record<string, string> = {
      undo: 'Undo',
      redo: 'Redo',
      'list-indent': 'Demote list item',
    };
    const rows = keybindingRows(bindings, (c) => titles[c]);
    expect(rows.map((r) => r.title)).toEqual([
      'Demote list item',
      'Redo',
      'save',
      'Undo',
    ]);
    expect(rows[1].keys).toEqual(['Shift-Mod-z', 'Mod-y']); // gathered
    expect(rows[2]).toMatchObject({ scope: 'window', source: 'desktop' });
    expect(
      filterKeybindingRows(rows, 'list', true).map((r) => r.command),
    ).toEqual(['list-indent']);
    expect(
      filterKeybindingRows(rows, '⌘Y', true).map((r) => r.command),
    ).toEqual(['redo']);
    expect(
      filterKeybindingRows(rows, 'desktop', true).map((r) => r.command),
    ).toEqual(['save']);
    expect(filterKeybindingRows(rows, '', true)).toHaveLength(4);
  });

  it('registry: keyed by the normalized chord, later add replaces, sources', () => {
    const r = new KeybindingRegistry(true);
    r.add({ key: 'Mod-Shift-z', command: 'redo', source: 'core' });
    r.add({ key: 'Shift-Mod-z', command: 'redo-2', source: 'host' }); // same chord
    expect(r.size).toBe(1);
    expect(r.get('Meta-Shift-z')?.command).toBe('redo-2'); // normalized lookup
    expect(r.get('Mod-Shift-z')?.source).toBe('host'); // raw lookup too
    r.add({ key: 'Mod-b', command: 'bold', source: 'core' });
    r.add({ key: 'Mod-s', command: 'save', scope: 'window', source: 'host' });
    expect(r.bySource('host').map((b) => b.command)).toEqual([
      'redo-2',
      'save',
    ]);
    let changes = 0;
    const off = r.onChange(() => changes++);
    r.clearSource('host');
    expect([...r].map((b) => b.command)).toEqual(['bold']);
    expect(changes).toBe(1);
    off();
    expect(() => r.add({ key: 'Mod+b', command: 'x', source: 's' })).toThrow();
    expect(() =>
      r.add({ key: 'Enter', command: 'x', scope: 'window', source: 's' }),
    ).toThrow(); // no modifier at window scope
  });
});
