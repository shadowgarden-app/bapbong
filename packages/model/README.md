# @shadow-garden/bapbong-model

The ProseMirror **document schema** for bapbong, plus list-numbering helpers.
This is the canonical in-memory document — `importDocx` produces it, the layout
engine consumes it, and the hidden editor edits it.

- **Scope:** `scope:model`
- **Depends on:** `prosemirror-model`

## What it provides

- **`schema`** — the main document `Schema`: `doc` (with undoable `attrs`
  `numbering` / `sections` / `comments`), `paragraph`, `text`, `image`, `table`,
  `hardBreak`, …; marks `strong`, `em`, `underline`, `strike`, `link`, `color`,
  `highlight`, `vertAlign`, `footnote { num }`, `comment { ids }`.
- **`commentSchema`** — a tiny schema for comment bodies: `doc` / `paragraph` /
  `text` + an inline **`mention`** atom (`{ id, label }`) with `leafText`/`toDOM`.
- **`createNumberingCounter(defs)`** — resolves live list markers (decimal,
  bullet, multilevel) from `doc.attrs.numbering`.
- Types: `BapbongSchema`, `CommentSchema`, `Align`.

## Build / test

```sh
pnpm nx build @shadow-garden/bapbong-model
pnpm nx test  @shadow-garden/bapbong-model
```
