import { deleteSelection } from 'prosemirror-commands';
import type { EditorState } from 'prosemirror-state';
import type { Command } from '@shadow-garden/bapbong-contracts';

/** Delete the current selection (no-op when the selection is empty). */
export function deleteSelectionCommand(): Command {
  return {
    name: 'delete',
    run: (state, dispatch) => deleteSelection(state, dispatch),
    isEnabled: (state) => !state.selection.empty,
  };
}

/**
 * Word's Alt+X: the hex digits before the caret become the character they
 * name (`2611` → ☑, `U+2611` too), and, when there are none, the character
 * before the caret becomes its code (☑ → `2611`) — the same key both ways. A
 * non-empty selection converts as a whole on the same terms. Marks of the
 * replaced text carry over (insertText inherits them), so a ☑ typed in bold
 * stays bold. Nothing to convert → not handled, and the key falls through.
 */
export function toggleUnicodeHex(): Command {
  return {
    name: 'toggle-unicode-hex',
    title: 'Toggle Unicode character ↔ hex code',
    run(state, dispatch) {
      const conv = unicodeHexConversion(state);
      if (!conv) return false;
      if (dispatch)
        dispatch(
          state.tr.insertText(conv.text, conv.from, conv.to).scrollIntoView(),
        );
      return true;
    },
    isEnabled: (state) => unicodeHexConversion(state) !== null,
  };
}

/** Longest hex run (2–6 digits, optional `U+` / `u+`) at the END of `s`
 *  that names a valid, non-control, non-surrogate code point. Prefers the
 *  longest so `1F600` wins over `F600`; falls back shorter when the long
 *  reading is past Unicode (`FFFFFF` → the last five). */
function trailingHexChar(s: string): { len: number; char: string } | null {
  const m = /(?:[Uu]\+)?([0-9A-Fa-f]{1,6})$/.exec(s);
  if (!m) return null;
  const digits = m[1];
  const prefix = m[0].length - digits.length;
  for (let n = Math.min(6, digits.length); n >= 2; n--) {
    const cp = parseInt(digits.slice(digits.length - n), 16);
    if (cp > 0x10ffff || (cp >= 0xd800 && cp <= 0xdfff)) continue;
    if (cp < 0x20 || (cp >= 0x7f && cp < 0xa0)) continue;
    // The prefix only counts when the whole digit run is taken.
    const len = n + (n === digits.length ? prefix : 0);
    return { len, char: String.fromCodePoint(cp) };
  }
  return null;
}

/** `U+XXXX`-style code of the last code point of `s` — the code Word writes
 *  is bare upper-case hex, at least four digits. */
function lastCharHex(s: string): { len: number; hex: string } | null {
  if (!s) return null;
  const cps = [...s];
  const ch = cps[cps.length - 1];
  const cp = ch.codePointAt(0) ?? 0;
  return {
    len: ch.length,
    hex: cp.toString(16).toUpperCase().padStart(4, '0'),
  };
}

function unicodeHexConversion(
  state: EditorState,
): { from: number; to: number; text: string } | null {
  const { from, to, empty, $from } = state.selection;
  if (!$from.parent.isTextblock) return null;
  if (!empty) {
    // A selection converts as a whole: hex → char, else char → hex.
    if ($from.parent !== state.selection.$to.parent) return null;
    const sel = state.doc.textBetween(from, to, '￼');
    const asHex = trailingHexChar(sel);
    if (asHex && asHex.len === sel.length)
      return { from, to, text: asHex.char };
    if ([...sel].length === 1) {
      const h = lastCharHex(sel);
      return h ? { from, to, text: h.hex } : null;
    }
    return null;
  }
  // Text before the caret in this textblock (leaf nodes as U+FFFC, which
  // is not hex, so an image ends the run).
  const before = state.doc.textBetween($from.start(), from, '￼', '￼');
  const hex = trailingHexChar(before);
  if (hex) return { from: from - hex.len, to: from, text: hex.char };
  const h = lastCharHex(before);
  if (!h || before.endsWith('￼')) return null;
  return { from: from - h.len, to: from, text: h.hex };
}
