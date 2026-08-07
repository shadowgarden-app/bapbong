import { Collection } from '@shadow-garden/bapbong-contracts';
import type { Command } from '@shadow-garden/bapbong-contracts';
import { toggleMarkCommand, clearMarks } from './marks.js';
import { setAlign, toggleHeading } from './paragraph.js';
import { toggleList } from './list.js';
import { redoCommand, undoCommand } from './history.js';
import { pageBreakCommand } from './insert.js';
import { insertSectionBreak, setColumns } from './sections.js';
import {
  insertLandscapeSection,
  setMargins,
  setOrientation,
  setPaperSize,
} from './page-setup.js';

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
      // Marks the importer has always produced but nothing could switch on.
      toggleMarkCommand('small-caps', 'smallCaps'),
      toggleMarkCommand('double-strike', 'dstrike'),
      clearMarks(),
      setAlign('left'),
      setAlign('center'),
      setAlign('right'),
      setAlign('justify'),
      toggleList('bullet'),
      toggleList('ordered'),
      toggleHeading(1),
      toggleHeading(2),
      toggleHeading(3),
      toggleHeading(4),
      toggleHeading(5),
      toggleHeading(6),
      undoCommand(),
      redoCommand(),
      pageBreakCommand(),
      insertSectionBreak({ newPage: true }),
      insertSectionBreak({ newPage: false }),
      setColumns(1),
      setColumns(2),
      setColumns(3),
      setOrientation('portrait'),
      setOrientation('landscape'),
      setPaperSize('a4'),
      setPaperSize('letter'),
      setPaperSize('legal'),
      setPaperSize('executive'),
      setPaperSize('a3'),
      setPaperSize('a5'),
      setMargins('normal'),
      setMargins('narrow'),
      setMargins('moderate'),
      setMargins('wide'),
      setMargins('office2003'),
      insertLandscapeSection(),
    ],
    { idProperty: 'name' },
  );
}
