import type { DocCompat } from '@shadow-garden/bapbong-contracts';
import {
  attrOf,
  child,
  children,
  isToggleOn,
  type OoxmlNode,
} from './ooxml.js';

/** The mode Word assumes for a document that names none: Word 2007. */
export const DEFAULT_COMPAT_MODE = 12;

/**
 * Resolve `word/settings.xml` (its `w:settings` element, or undefined when the
 * part is missing) into the document's {@link DocCompat}. Pure: every field is
 * derived here, once, and consumers only ever read the resolved answers.
 *
 * Reads exactly what it resolves. A `w:compatSetting` is found by NAME (so
 * every entry's name is read) but its VALUE is read only for the settings a
 * field consumes — an unread `@w:val` in the XML audit is the honest sign of
 * a compat setting this program has not adopted yet.
 */
export function parseCompat(settings: OoxmlNode | undefined): DocCompat {
  const compat = child(settings, 'w:compat');
  const setting = (name: string): string | undefined =>
    attrOf(
      children(compat, 'w:compatSetting').find(
        (c) => attrOf(c, 'w:name') === name,
      ),
      'w:val',
    );
  const parsed = Number(setting('compatibilityMode'));
  const mode =
    Number.isFinite(parsed) && parsed > 0 ? parsed : DEFAULT_COMPAT_MODE;
  // CT_Compat is where this toggle lives (ECMA-376 §17.15.3.24); a settings-
  // level placement is tolerated for documents that put it there.
  const noHtmlAutoSpacing =
    child(compat, 'w:doNotUseHTMLParagraphAutoSpacing') ??
    child(settings, 'w:doNotUseHTMLParagraphAutoSpacing');
  const overrideVal = setting('overrideTableStyleFontSizeAndJustification');
  return {
    mode,
    htmlAutoSpacing: !isToggleOn(noHtmlAutoSpacing),
    tableIndentToBorder: mode >= 15,
    // MS-DOCX 2.6.x: absent/false is the Word 2007 evaluation, true is the
    // ISO §17.7.2 hierarchy (which is the cascade the importer builds).
    normalStyleYieldsToTableStyle: !(
      overrideVal !== undefined && !['0', 'false', 'off'].includes(overrideVal)
    ),
    underlineTrailingSpaces: isToggleOn(child(compat, 'w:ulTrailSpace')),
    expandLineBeforeSoftBreak: !isToggleOn(
      child(compat, 'w:doNotExpandShiftReturn'),
    ),
  };
}
