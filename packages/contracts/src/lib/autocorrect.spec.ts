import { AUTOCORRECT_RULES, autoCorrectMatch } from './autocorrect.js';

describe("autoCorrect (Word's symbol rules)", () => {
  it('matches a sequence only as its last character is typed', () => {
    expect(autoCorrectMatch('(c', ')')).toEqual({ seq: '(c)', to: '©' });
    expect(autoCorrectMatch('see (tm', ')')?.to).toBe('™');
    expect(autoCorrectMatch('a --', '>')?.to).toBe('→');
    expect(autoCorrectMatch('..', '.')?.to).toBe('…');
    // The typed character alone never completes a rule.
    expect(autoCorrectMatch('', ')')).toBeNull();
    // Nothing partial: "(c" is not "(c)".
    expect(autoCorrectMatch('(', 'c')).toBeNull();
    // Every rule's sequence is at least two characters long, so a single
    // keystroke can never be a whole rule.
    for (const r of AUTOCORRECT_RULES) expect(r.seq.length).toBeGreaterThan(1);
  });
});
