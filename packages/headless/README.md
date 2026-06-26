# @shadow-garden/bapbong-headless

The **backend façade** for bapbong. One install that re-exports the entire
*isomorphic tier* — no DOM, no canvas, no `prosemirror-view` — so a server can
read, edit, and write `.docx` with zero browser dependencies.

It re-exports:

| Package | What you get |
| --- | --- |
| `@shadow-garden/bapbong-contracts` | Shared types — `PageConfig`, `Align`, `Command`, `Collection`, … |
| `@shadow-garden/bapbong-model` | The ProseMirror `schema` (+ `numbering`) |
| `@shadow-garden/bapbong-docx` | `importDocx` / `exportDocx` round-trip |
| `@shadow-garden/bapbong-commands` | Headless editor ops — `toggleMarkCommand`, `setAlign`, `insertRow`, `defaultCommands()`, … |

## Install

```sh
pnpm add @shadow-garden/bapbong-headless prosemirror-state
```

`prosemirror-state` is a peer of the workflow below — the commands operate on an
`EditorState`, which you construct yourself (the same one the editor UI uses).

## Backend round-trip

```ts
import {
  importDocx,
  exportDocx,
  toggleMarkCommand,
  setAlign,
} from '@shadow-garden/bapbong-headless';
import { EditorState, TextSelection } from 'prosemirror-state';

// 1. Read a .docx (ArrayBuffer | Uint8Array | Blob).
const imported = await importDocx(bytes);

// 2. Drive the SAME commands the editor UI uses, on a Node EditorState.
let state = EditorState.create({ doc: imported.doc });
state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 12)));
toggleMarkCommand('bold', 'strong').run(state, (tr) => (state = state.apply(tr)));
setAlign('center').run(state, (tr) => (state = state.apply(tr)));

// 3. Write it back. `carry` keeps parts bapbong doesn't model yet (styles,
//    numbering, headers/footers, …) instead of dropping them.
const out = await exportDocx(state.doc, { carry: imported.raw });
```

## Boundary

This package is tagged `scope:headless` and lint-guarded against DOM globals
(`document`, `window`, `Image`, `DOMParser`, …). **Do not** add
`bapbong-editor` or `bapbong-view` to it — they pull in canvas/DOM and would
break Node usage. For an in-browser preview use `@shadow-garden/bapbong-view`;
for a full editor use `@shadow-garden/bapbong-editor`.
