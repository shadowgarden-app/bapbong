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
/** Human titles for the built-ins — what the Keyboard-shortcuts dialog and
 *  other listings show. Names without an entry fall back to the name. */
const TITLES: Record<string, string> = {
  bold: 'Bold',
  italic: 'Italic',
  underline: 'Underline',
  strike: 'Strikethrough',
  superscript: 'Superscript',
  subscript: 'Subscript',
  'small-caps': 'Small caps',
  'double-strike': 'Double strikethrough',
  'clear-format': 'Clear formatting',
  'align-left': 'Align left',
  'align-center': 'Center',
  'align-right': 'Align right',
  'align-justify': 'Justify',
  'bullet-list': 'Bullet list',
  'ordered-list': 'Numbered list',
  'heading-1': 'Heading 1',
  'heading-2': 'Heading 2',
  'heading-3': 'Heading 3',
  'heading-4': 'Heading 4',
  'heading-5': 'Heading 5',
  'heading-6': 'Heading 6',
  undo: 'Undo',
  redo: 'Redo',
  'page-break': 'Page break before',
  'section-break-next-page': 'Section break (next page)',
  'section-break-continuous': 'Section break (continuous)',
  'columns-1': 'One column',
  'columns-2': 'Two columns',
  'columns-3': 'Three columns',
  'orientation-portrait': 'Portrait',
  'orientation-landscape': 'Landscape',
  'paper-a4': 'Paper: A4',
  'paper-letter': 'Paper: Letter',
  'paper-legal': 'Paper: Legal',
  'paper-executive': 'Paper: Executive',
  'paper-a3': 'Paper: A3',
  'paper-a5': 'Paper: A5',
  'margins-normal': 'Margins: Normal',
  'margins-narrow': 'Margins: Narrow',
  'margins-moderate': 'Margins: Moderate',
  'margins-wide': 'Margins: Wide',
  'margins-office2003': 'Margins: Office 2003',
  'insert-landscape-section': 'Insert landscape section',
};

/** `cmd` with its title from {@link TITLES} — a spread keeps the command's
 *  own methods (`run`/`isActive`/`isEnabled` are plain properties here). */
function titled(cmd: Command): Command {
  const title = TITLES[cmd.name];
  return title ? { ...cmd, title } : cmd;
}

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
    ].map(titled),
    { idProperty: 'name' },
  );
}
