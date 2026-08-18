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
 */
export function parseCompat(settings: OoxmlNode | undefined): DocCompat {
  const compat = child(settings, 'w:compat');
  const modeAttr = attrOf(
    children(compat, 'w:compatSetting').find(
      (c) => attrOf(c, 'w:name') === 'compatibilityMode',
    ),
    'w:val',
  );
  const parsed = Number(modeAttr);
  const mode =
    Number.isFinite(parsed) && parsed > 0 ? parsed : DEFAULT_COMPAT_MODE;
  return {
    mode,
    htmlAutoSpacing: !isToggleOn(
      child(settings, 'w:doNotUseHTMLParagraphAutoSpacing'),
    ),
    tableIndentToBorder: mode >= 15,
  };
}
