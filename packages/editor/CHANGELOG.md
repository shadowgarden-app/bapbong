## 0.1.0 (2026-07-15)

### 🚀 Features

- **editor:** onKey plugin hook; Escape cancels image gestures (M16-V2) ([4b06811](https://github.com/shadowgarden-app/bapbong/commit/4b06811))
- **editor:** hover cursors for image resize/rotate + verify (M16-R6) ([4917aa5](https://github.com/shadowgarden-app/bapbong/commit/4917aa5))
- **editor:** rotate images via the frame knob (M16-R5) ([f47be50](https://github.com/shadowgarden-app/bapbong/commit/f47be50))
- **editor:** drag-resize images, committing once on release (M16-R3) ([d5bacd0](https://github.com/shadowgarden-app/bapbong/commit/d5bacd0))
- **editor:** select images with a resize frame (M16-R2) ([abc58cf](https://github.com/shadowgarden-app/bapbong/commit/abc58cf))
- **editor:** wire font-metrics measurer as default + canvas fallback ([#4](https://github.com/shadowgarden-app/bapbong/pull/4))
- **editor:** expose print + File ▸ Print (Group 3 #2) ([#2](https://github.com/shadowgarden-app/bapbong/issues/2))
- **editor:** expose zoom + zoom dropdown in the playground toolbar (Group 3 #1) ([#1](https://github.com/shadowgarden-app/bapbong/issues/1))
- **sections:** in-document section-break markers + click-to-delete (B3.6) ([f66679f](https://github.com/shadowgarden-app/bapbong/commit/f66679f))
- **a11y:** ARIA shadow-DOM mirror + Ctrl/Cmd-click hyperlinks (M10) ([ca353cc](https://github.com/shadowgarden-app/bapbong/commit/ca353cc))
- **view:** extract RenderCore + BapbongView; editor composes it (M9) ([78f956e](https://github.com/shadowgarden-app/bapbong/commit/78f956e))
- **ui:** cell-properties Borders panel — 10 presets + width/style/colour ([23f5574](https://github.com/shadowgarden-app/bapbong/commit/23f5574))
- cell-properties dialog applied to selected cells (icon + right-click) ([6030a4e](https://github.com/shadowgarden-app/bapbong/commit/6030a4e))
- **editor:** cell-selection action icon + editor.tableSelection handle ([b78dc1e](https://github.com/shadowgarden-app/bapbong/commit/b78dc1e))
- **editor:** cell-range selection (drag across table cells) ([a63b2c6](https://github.com/shadowgarden-app/bapbong/commit/a63b2c6))
- right-click context menu (clipboard defaults + table ops in a cell) ([ece783e](https://github.com/shadowgarden-app/bapbong/commit/ece783e))
- **editor:** column resize keeps table width fixed; drag shows a guide, commits on drop ([ceec68c](https://github.com/shadowgarden-app/bapbong/commit/ceec68c))
- **editor:** L2 pointer hook + table-column resize plugin ([bf17d87](https://github.com/shadowgarden-app/bapbong/commit/bf17d87))
- **editor:** expose editor.commands registry (Collection<Command>) ([83ddaf9](https://github.com/shadowgarden-app/bapbong/commit/83ddaf9))
- **contracts:** Collection takes its id property via the constructor ([ba8167b](https://github.com/shadowgarden-app/bapbong/commit/ba8167b))
- **editor:** compose document schema from plugin contributions (P4b) ([3fe06f6](https://github.com/shadowgarden-app/bapbong/commit/3fe06f6))
- **editor:** generic plugin paint decorations (additive) ([8b2fbdb](https://github.com/shadowgarden-app/bapbong/commit/8b2fbdb))
- **editor:** add EditorPlugin contract + plugin lifecycle/event fan-out ([e727085](https://github.com/shadowgarden-app/bapbong/commit/e727085))
- **editor:** extract framework-agnostic BapbongEditor core ([276cab4](https://github.com/shadowgarden-app/bapbong/commit/276cab4))
- **docx,model:** M1 — DOCX import to ProseMirror model ([07d236b](https://github.com/shadowgarden-app/bapbong/commit/07d236b))

### 🩹 Fixes

- **editor:** right-click no longer collapses the selection ([b771915](https://github.com/shadowgarden-app/bapbong/commit/b771915))
- **editor:** guard setPointerCapture so a throw can't abort caret placement ([9dd1bbb](https://github.com/shadowgarden-app/bapbong/commit/9dd1bbb))

### 🧱 Updated Dependencies

- Updated @shadow-garden/bapbong-input-bridge to 0.1.0
- Updated @shadow-garden/bapbong-contracts to 0.1.0
- Updated @shadow-garden/bapbong-selection to 0.1.0
- Updated @shadow-garden/bapbong-commands to 0.1.0
- Updated @shadow-garden/bapbong-model to 0.1.0
- Updated @shadow-garden/bapbong-view to 0.1.0

### ❤️ Thank You

- Le Phuoc Minh