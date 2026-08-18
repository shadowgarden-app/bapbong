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
        normalStyleYieldsToTableStyle: true,
        underlineTrailingSpaces: false,
        expandLineBeforeSoftBreak: true,
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
    // Where Word writes it — inside w:compat — and, tolerated, at the
    // settings level.
    expect(
      parseCompat(
        settings('<w:compat><w:doNotUseHTMLParagraphAutoSpacing/></w:compat>'),
      ).htmlAutoSpacing,
    ).toBe(false);
    expect(
      parseCompat(settings('<w:doNotUseHTMLParagraphAutoSpacing/>'))
        .htmlAutoSpacing,
    ).toBe(false);
    expect(
      parseCompat(
        settings(
          '<w:compat><w:doNotUseHTMLParagraphAutoSpacing w:val="0"/></w:compat>',
        ),
      ).htmlAutoSpacing,
    ).toBe(true);
  });

  it('reads the two line-level toggles: ulTrailSpace, doNotExpandShiftReturn', () => {
    const on = parseCompat(
      settings(
        '<w:compat><w:ulTrailSpace/><w:doNotExpandShiftReturn/></w:compat>',
      ),
    );
    expect(on).toMatchObject({
      underlineTrailingSpaces: true,
      expandLineBeforeSoftBreak: false,
    });
    const off = parseCompat(
      settings(
        '<w:compat><w:ulTrailSpace w:val="0"/><w:doNotExpandShiftReturn w:val="false"/></w:compat>',
      ),
    );
    expect(off).toMatchObject({
      underlineTrailingSpaces: false,
      expandLineBeforeSoftBreak: true,
    });
  });

  it('overrideTableStyleFontSizeAndJustification: absent/0 is the 2007 reading', () => {
    const cs = (val: string) =>
      `<w:compat><w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="${val}"/></w:compat>`;
    expect(parseCompat(undefined).normalStyleYieldsToTableStyle).toBe(true);
    expect(parseCompat(settings(cs('0'))).normalStyleYieldsToTableStyle).toBe(
      true,
    );
    expect(parseCompat(settings(cs('1'))).normalStyleYieldsToTableStyle).toBe(
      false,
    );
  });
});
