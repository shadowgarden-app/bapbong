import {
  SPECIAL_CHARACTERS,
  SYMBOL_GROUPS,
  SYMBOL_NAMES,
  codePointLabel,
  parseCodePoint,
  pushRecent,
  searchSymbols,
} from './symbol-sets.js';

describe('symbol sets', () => {
  it('every entry is one real, non-PUA code point with a name', () => {
    for (const grp of SYMBOL_GROUPS)
      for (const e of grp.entries) {
        expect([...e.char]).toHaveLength(1); // one code point, astral allowed
        const cp = e.char.codePointAt(0)!;
        expect(cp >= 0xe000 && cp <= 0xf8ff).toBe(false); // no symbol-font PUA
        expect(e.name.length).toBeGreaterThan(1); // "Mu", "Nu", "Xi", "Pi"
      }
    // Group ids are stable and unique (recents / tests key on them).
    const ids = SYMBOL_GROUPS.map((g) => g.id);
    expect(new Set(ids).size).toBe(ids.length);
    // The names index reaches into the special-characters tab too.
    expect(SYMBOL_NAMES.get('☑')).toBe('Ballot box with check');
    expect(SYMBOL_NAMES.get(' ')).toBe('Nonbreaking space');
    expect(SPECIAL_CHARACTERS.some((c) => c.char === '—')).toBe(true);
  });

  it('labels and parses code points the way the dialog shows them', () => {
    expect(codePointLabel('☑')).toBe('U+2611');
    expect(codePointLabel('A')).toBe('U+0041');
    expect(codePointLabel('😀')).toBe('U+1F600');
    for (const s of ['2611', 'U+2611', 'u+2611', '0x2611', ' 2611 '])
      expect(parseCodePoint(s)).toBe('☑');
    expect(parseCodePoint('1F600')).toBe('😀');
    expect(parseCodePoint('zz')).toBeNull();
    expect(parseCodePoint('D800')).toBeNull(); // a lone surrogate
    expect(parseCodePoint('110000')).toBeNull(); // past Unicode
    expect(parseCodePoint('0007')).toBeNull(); // a control character
    expect(parseCodePoint('')).toBeNull();
  });

  it('keeps a most-recent-first row of sixteen with no repeats', () => {
    let r: string[] = [];
    for (const c of ['a', 'b', 'c']) r = pushRecent(r, c);
    expect(r).toEqual(['c', 'b', 'a']);
    expect(pushRecent(r, 'a')).toEqual(['a', 'c', 'b']); // moves to the front
    let many: string[] = [];
    for (let i = 0; i < 20; i++) many = pushRecent(many, String(i));
    expect(many).toHaveLength(16);
    expect(many[0]).toBe('19');
  });

  it('searches names across groups, in grid order, without duplicates', () => {
    const r = searchSymbols('ballot');
    expect(r.map((e) => e.char)).toEqual(['☐', '☑', '☒', '✗', '✘']);
    // A group whose label matches leads: "arrow" → the Arrows group before
    // the arrowhead bullets that merely contain the word.
    expect(searchSymbols('arrow')[0].char).toBe('←');
    expect(searchSymbols('arrow').some((e) => e.char === '➢')).toBe(true);
    // "Check mark" sits in two groups; it comes back once.
    expect(searchSymbols('check mark').map((e) => e.char)).toEqual(['✓', '✔']);
    expect(searchSymbols('   ')).toEqual([]);
  });
});
