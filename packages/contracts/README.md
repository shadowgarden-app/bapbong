# @shadow-garden/bapbong-contracts

The shared **vocabulary** every other bapbong package speaks: paint-ready types,
the plugin contract, and one small isomorphic utility (`Collection`). Almost all
types; the only runtime code is `Collection`.

- **Scope:** `scope:pure` (leaf — imports no other workspace package)
- **Depends on:** `prosemirror-state`, `prosemirror-model` — **type-only**, for the
  plugin/editor-change contracts. No DOM: the package is **isomorphic** (runs on
  Node/server; enforced by a `no-restricted-globals` lint guard).

## What it defines

- **Flow model** (layout input): `FlowParagraph`, `FlowTable` /
  `FlowTableRow` / `FlowTableCell`, `FlowFloat`, and inlines
  `InlineRun` / `InlineImage` / `InlineBreak` / `InlineField`.
- **Resolved layout** (paint-ready engine output): `ResolvedLayout`,
  `ResolvedPage`, `LayoutLine`, `LayoutSegment`, `LayoutImageSegment`,
  `ResolvedCell`, `ResolvedFloat`, `ResolvedFootnotes`, `ResolvedChrome`.
- **Geometry & config:** `PageConfig`, `LayoutConfig`, `ColumnConfig`,
  `ParagraphSpacing`, `ParagraphIndent`, `CellPadding`, `CaretRect`,
  `PagePoint`, `SelectionRect`.
- **Fonts / measurement:** `FontSpec`, `FontMetrics`, `MeasureText`, `MeasureMetrics`.
- **Comments:** `IUser`, `CommentNode`, `CommentData`.
- **Plugin contract** (the stable surface the editor + plugins share, so neither
  imports the other): `EditorPlugin`, `PluginContext`, `EditorChange`,
  `RangeDecoration`, `PaintDecoration`.
- **Command contract**: `Command` (`{ name, run, isActive?, isEnabled? }`) +
  `Dispatch` — the headless editor-operation surface the toolbar/menubar,
  plugins and a backend share. Implementations live in
  `@shadow-garden/bapbong-commands`.
- **`Collection<T>`** — a name/id-keyed, insertion-ordered collection
  (`get`/`add`/`remove`/`has`/`entries` + iterable). Keys by `id` by default, or
  `new Collection(items, { idProperty })`; throws if an item lacks its key.
  The editor uses it as its plugin registry.

## Build / test

```sh
pnpm nx build @shadow-garden/bapbong-contracts
pnpm nx test  @shadow-garden/bapbong-contracts
```
