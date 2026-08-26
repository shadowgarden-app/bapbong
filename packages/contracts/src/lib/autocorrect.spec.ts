import { AUTOCORRECT_RULES, autoCorrectMatch } from './autocorrect.js';
import { mathAutoCorrectMatch } from './math.js';

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

describe("mathAutoCorrectMatch (Word's Math AutoCorrect)", () => {
  it('completes \\name at the end of the typed text', () => {
    expect(mathAutoCorrectMatch('Z = \\omega')).toEqual({
      length: 6,
      to: 'ω',
    });
    expect(mathAutoCorrectMatch('\\Delta')?.to).toBe('Δ');
    expect(mathAutoCorrectMatch('x \\in')?.to).toBe('∈');
    expect(mathAutoCorrectMatch('m \\doubleZ')?.to).toBe('ℤ');
  });

  it('keeps Word-faithful variant names', () => {
    expect(mathAutoCorrectMatch('\\epsilon')?.to).toBe('ϵ');
    expect(mathAutoCorrectMatch('\\varepsilon')?.to).toBe('ε');
    expect(mathAutoCorrectMatch('\\phi')?.to).toBe('ϕ');
    expect(mathAutoCorrectMatch('\\varphi')?.to).toBe('φ');
  });

  it('ignores unknown names, missing backslashes, and dead prefixes', () => {
    expect(mathAutoCorrectMatch('\\banana')).toBeNull();
    expect(mathAutoCorrectMatch('omega')).toBeNull();
    expect(mathAutoCorrectMatch('\\omega2')).toBeNull();
    expect(mathAutoCorrectMatch('')).toBeNull();
  });
});
