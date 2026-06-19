# @shadow-garden/bapbong-contracts

Shared **types only** for the bapbong pipeline — the vocabulary every other
package speaks. No runtime code.

- **Scope:** `scope:pure` (leaf — may not import any other workspace package)
- **Depends on:** nothing

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

## Build / test

```sh
pnpm nx build @shadow-garden/bapbong-contracts
pnpm nx test  @shadow-garden/bapbong-contracts
```
