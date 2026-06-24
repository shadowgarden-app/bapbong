# @shadow-garden/bapbong-commands

**Headless editor operations** — the L1 layer that both the UI (menubar /
toolbar / context menu / dialogs) and a **backend** drive. A command operates
purely on a ProseMirror `EditorState` with no DOM, so the *same* code edits a
document in the browser and on the server.

- **Scope:** `scope:pure` (isomorphic — depends only on `contracts`; runs on
  Node, enforced by a `no-restricted-globals` lint guard)
- **Depends on:** `@shadow-garden/bapbong-contracts` (for `Command` + `Collection`),
  `prosemirror-state`, `prosemirror-model`, `prosemirror-commands`

## The shape

A `Command` (defined in `contracts`) is `{ name, run(state, dispatch?), isActive?, isEnabled? }`
— ProseMirror's command convention plus metadata for menus. Omit `dispatch` to
probe whether it would apply.

- **Marks** (`marks.ts`): `toggleMarkCommand(id, markName?)`, `isMarkActive`.
- **Paragraph** (`paragraph.ts`): `setAlign(align)`, `activeAlign`, `ALIGNMENTS`.
- **Table** (`table.ts`): `cellAt` (query), `setCellAttrs(pos, attrs)`,
  `setCellBackground(color)`, `setColumnWidth(cellPos, width)`.
- **Registry** (`registry.ts`): `defaultCommands(): Collection<Command>` — the
  built-in static commands a toolbar renders by name. Parameterized ops (a
  picked colour, a dragged width) are called via their factory functions.

## Backend usage

No editor, no canvas, no `prosemirror-view`. Build a state and run commands:

```ts
import { EditorState } from 'prosemirror-state';
import { defaultCommands, setCellBackground } from '@shadow-garden/bapbong-commands';

const state = EditorState.create({ schema, doc });          // headless, Node
const next = applyTr(state, (s, d) => defaultCommands().get('bold')!.run(s, d));
applyTr(next, (s, d) => setCellBackground('#ffd600').run(s, d));
// → export the resulting doc back to .docx
```

## Build / test

```sh
pnpm nx build @shadow-garden/bapbong-commands
pnpm nx test  @shadow-garden/bapbong-commands   # runs in Node (no DOM)
```
