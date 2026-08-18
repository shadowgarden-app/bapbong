import { DEFAULT_COMPAT_MODE, parseCompat } from './compat.js';
import { child, parseXml } from './ooxml.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const settings = (inner: string) =>
  child(
    parseXml(
      `<?xml version="1.0"?><w:settings xmlns:w="${W_NS}">${inner}</w:settings>`,
    ),
    'w:settings',
  );
const mode = (n: number | string) =>
  `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${n}"/></w:compat>`;

describe('parseCompat', () => {
  it('reads a document naming no mode as Word 2007 (12), current-Word rules off', () => {
    // No settings part at all, and a settings part without w:compat, are the
    // same document to Word: compatibilityMode 12.
    for (const el of [undefined, settings('')]) {
      expect(parseCompat(el)).toEqual({
        mode: DEFAULT_COMPAT_MODE,
        htmlAutoSpacing: true,
        tableIndentToBorder: false,
      });
    }
    expect(DEFAULT_COMPAT_MODE).toBe(12);
  });

  it('resolves the mode into the rules that depend on it', () => {
    expect(parseCompat(settings(mode(14)))).toMatchObject({
      mode: 14,
      tableIndentToBorder: false, // Word 2010: tblInd is to the cell text
    });
    expect(parseCompat(settings(mode(15)))).toMatchObject({
      mode: 15,
      tableIndentToBorder: true, // Word 2013+: to the border
    });
    // A mode Word has never written (a garbage value) falls back to 12
    // rather than to NaN comparisons that would silently pick one branch.
    expect(parseCompat(settings(mode('x'))).mode).toBe(12);
  });

  it('reads the individual compat toggles, w:val aware', () => {
    expect(
      parseCompat(settings('<w:doNotUseHTMLParagraphAutoSpacing/>'))
        .htmlAutoSpacing,
    ).toBe(false);
    expect(
      parseCompat(settings('<w:doNotUseHTMLParagraphAutoSpacing w:val="0"/>'))
        .htmlAutoSpacing,
    ).toBe(true);
  });
});
