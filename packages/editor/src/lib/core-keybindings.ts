import type { Keybinding } from '@shadow-garden/bapbong-contracts';

/**
 * The editor's own chords, as data. The editor registers exactly these at
 * construction; anything that needs to KNOW them without an editor — the
 * desktop app's generated UI guide, written at ship time with no window —
 * reads this table instead of reaching into a live instance.
 *
 * Pure on purpose: importing it must not touch the DOM.
 */
export const CORE_KEYBINDINGS: readonly Keybinding[] = [
  // Order matters only for the dialog's insertion order (it sorts anyway).
  { key: 'Enter', command: 'paragraph-enter', when: 'editing text' },
  {
    key: 'Backspace',
    command: 'backspace-outdent',
    when: 'at the start of a list or indented paragraph',
  },
  { key: 'Tab', command: 'list-indent', when: 'in a list' },
  { key: 'Shift-Tab', command: 'list-outdent', when: 'in a list' },
  { key: 'ArrowUp', command: 'caret-up', when: 'editing text' },
  { key: 'ArrowDown', command: 'caret-down', when: 'editing text' },
  { key: 'Shift-ArrowUp', command: 'select-up', when: 'editing text' },
  { key: 'Shift-ArrowDown', command: 'select-down', when: 'editing text' },
  { key: 'Mod-z', command: 'undo' },
  { key: 'Shift-Mod-z', command: 'redo' },
  { key: 'Mod-y', command: 'redo' },
  // Word's staples. Nothing bound these before; the base keymap has none.
  { key: 'Mod-b', command: 'bold', when: 'editing text' },
  { key: 'Mod-i', command: 'italic', when: 'editing text' },
  { key: 'Mod-u', command: 'underline', when: 'editing text' },
  // Word's Alt+X: hex before the caret ↔ the character.
  {
    key: 'Alt-x',
    command: 'toggle-unicode-hex',
    when: 'hex digits or a character before the caret',
  },
  // Word's Alt+=: insert (or convert the selection into) an equation.
  { key: 'Alt-=', command: 'insert-equation', when: 'editing text' },
].map((b) => ({ ...b, source: 'core' }));
