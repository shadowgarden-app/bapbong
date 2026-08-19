/**
 * Word's AutoCorrect, the symbol part: a typed sequence becomes the character
 * it stands for the moment it completes — `(c)` → ©, `-->` → →, `...` → …
 * — replacing the sequence in one transaction that keeps the run's marks. The
 * list is Word's default set for symbols (not its spelling fixes, and not the
 * smart quotes, which are locale work of their own). Applied on the typed
 * character, so a sequence a user WANTS literally is one ⌘Z away (the undo
 * restores what was typed).
 */
export const AUTOCORRECT_RULES: readonly { seq: string; to: string }[] = [
  { seq: '(c)', to: '©' },
  { seq: '(C)', to: '©' },
  { seq: '(r)', to: '®' },
  { seq: '(R)', to: '®' },
  { seq: '(tm)', to: '™' },
  { seq: '(TM)', to: '™' },
  { seq: '(e)', to: '€' },
  { seq: '...', to: '…' },
  { seq: '-->', to: '→' },
  { seq: '<--', to: '←' },
  { seq: '<=>', to: '⇔' },
  { seq: '==>', to: '⇒' },
  { seq: '<==', to: '⇐' },
];

/** The rule a typed `text` completes, given the `before` text — or null. */
export function autoCorrectMatch(
  before: string,
  text: string,
): { seq: string; to: string } | null {
  const joined = before + text;
  for (const r of AUTOCORRECT_RULES)
    if (joined.endsWith(r.seq) && r.seq.length > text.length) return r;
  return null;
}
