import { Collection } from '@shadow-garden/bapbong-contracts';
import type { Command } from '@shadow-garden/bapbong-contracts';
import { toggleMarkCommand } from './marks.js';
import { setAlign } from './paragraph.js';
import { toggleList } from './list.js';
import { redoCommand, undoCommand } from './history.js';
import { pageBreakCommand } from './insert.js';

/**
 * The built-in static commands a toolbar/menubar references by name — mark
 * toggles, paragraph alignment, super/subscript, history and page break. The
 * UI picks which to surface (the catalog can grow without bloating any one
 * toolbar). Parameterized ops that need a runtime value (cell background,
 * column width, a colour, a table size, an image src, a link href) are called
 * via their factory functions directly, not from this registry.
 *
 * Returns a fresh {@link Collection} each call so a host can extend it (e.g. add
 * plugin-contributed commands) without mutating a shared instance.
 */
export function defaultCommands(): Collection<Command> {
  return new Collection<Command>(
    [
      toggleMarkCommand('bold', 'strong'),
      toggleMarkCommand('italic', 'em'),
      toggleMarkCommand('underline'),
      toggleMarkCommand('strike'),
      toggleMarkCommand('superscript', 'vertAlign', { value: 'super' }),
      toggleMarkCommand('subscript', 'vertAlign', { value: 'sub' }),
      setAlign('left'),
      setAlign('center'),
      setAlign('right'),
      setAlign('justify'),
      toggleList('bullet'),
      toggleList('ordered'),
      undoCommand(),
      redoCommand(),
      pageBreakCommand(),
    ],
    { idProperty: 'name' },
  );
}
