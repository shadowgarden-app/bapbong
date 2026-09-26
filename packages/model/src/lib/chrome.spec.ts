import { describe, expect, it } from 'vitest';
import { effectiveSectionChrome, type ImportedChrome } from './chrome.js';

const flat: ImportedChrome<string> = {
  sectionChrome: null,
  headers: { default: 'H' },
  footers: { default: 'F' },
  titlePg: false,
};
const doc = (attrs: Record<string, unknown>) => ({ attrs });
const parse = (json: unknown) =>
  typeof json === 'string' && json !== 'bad' ? `parsed:${json}` : null;

describe('effectiveSectionChrome', () => {
  it('without overrides is what the importer read', () => {
    expect(effectiveSectionChrome(doc({}), flat, parse)).toBeNull();
    const per = {
      ...flat,
      sectionChrome: [{ headers: {}, footers: {}, titlePg: true }],
    };
    expect(effectiveSectionChrome(doc({}), per, parse)).toBe(per.sectionChrome);
  });

  it('an override spreads a flat chrome over the sections and replaces one story', () => {
    const out = effectiveSectionChrome(
      doc({
        sections: [{}, {}],
        sectionChromeOverrides: {
          '1': { footers: { default: 'F2' }, headers: { first: 'bad' } },
        },
      }),
      flat,
      parse,
    );
    expect(out).toEqual([
      { headers: { default: 'H' }, footers: { default: 'F' }, titlePg: false },
      {
        headers: { default: 'H' },
        footers: { default: 'parsed:F2' },
        titlePg: false,
      },
    ]);
  });
});
