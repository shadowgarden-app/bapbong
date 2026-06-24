import { Collection } from '@shadow-garden/bapbong-contracts';
import type { Command } from '@shadow-garden/bapbong-contracts';
import { toggleMarkCommand } from './marks.js';
import { setAlign } from './paragraph.js';

/**
 * The built-in static commands a toolbar/menubar renders by default — mark
 * toggles + paragraph alignment, keyed by `name`. Parameterized ops that need a
 * runtime value (cell background/colour, column width) are called via their
 * factory functions directly, not from this registry.
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
      setAlign('left'),
      setAlign('center'),
      setAlign('right'),
      setAlign('justify'),
    ],
    { idProperty: 'name' },
  );
}
