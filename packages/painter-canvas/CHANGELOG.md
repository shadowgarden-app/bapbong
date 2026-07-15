## 0.1.0 (2026-07-15)

### 🚀 Features

- **layout,painter:** contain rotated images vertically, clip horizontally (M16-V2) ([491656c](https://github.com/shadowgarden-app/bapbong/commit/491656c))
- **model,layout,painter,docx:** image rotation pipeline (M16-R4) ([77573a7](https://github.com/shadowgarden-app/bapbong/commit/77573a7))
- **painter,docx:** draw ellipse, roundRect, rightArrow, horizontalScroll ([33f4b72](https://github.com/shadowgarden-app/bapbong/commit/33f4b72))
- **layout,docx,painter:** render text inside wps:txbx textboxes ([2c25141](https://github.com/shadowgarden-app/bapbong/commit/2c25141))
- **layout,painter:** anchored floats position inside table cells (M14-P2b) ([4bf0aff](https://github.com/shadowgarden-app/bapbong/commit/4bf0aff))
- **docx,layout,painter:** drawn shapes — wps rect/line render + round-trip (M14-P2) ([212489f](https://github.com/shadowgarden-app/bapbong/commit/212489f))
- table borders carry width/style/colour (model + import/export + painter) ([e9bc931](https://github.com/shadowgarden-app/bapbong/commit/e9bc931))
- **editor:** generic plugin paint decorations (additive) ([8b2fbdb](https://github.com/shadowgarden-app/bapbong/commit/8b2fbdb))
- **comments:** authoring — add/reply/resolve/delete/edit (Phase A) ([3372c94](https://github.com/shadowgarden-app/bapbong/commit/3372c94))
- **docx,painter-canvas,playground:** comments — import + sidebar (read-only) ([bccb803](https://github.com/shadowgarden-app/bapbong/commit/bccb803))
- **layout-engine,docx,painter-canvas:** first/even header & footer variants ([58a51f3](https://github.com/shadowgarden-app/bapbong/commit/58a51f3))
- **layout-engine,docx,painter-canvas:** page-bottom footnote layout ([3314dc4](https://github.com/shadowgarden-app/bapbong/commit/3314dc4))
- **docx,model,layout-engine,painter-canvas:** T2 — per-cell borders + symbols ([928d44d](https://github.com/shadowgarden-app/bapbong/commit/928d44d))
- **model,docx,layout-engine,painter-canvas:** T2 — superscript / subscript ([3c227a8](https://github.com/shadowgarden-app/bapbong/commit/3c227a8))
- **model,docx,layout-engine,painter-canvas:** T1 — shading & highlight ([6149604](https://github.com/shadowgarden-app/bapbong/commit/6149604))
- **layout-engine,docx,model,painter-canvas:** M5 — floating images with text wrap ([bcab713](https://github.com/shadowgarden-app/bapbong/commit/bcab713))
- **docx,model,layout-engine,painter-canvas:** M5 — live PAGE/NUMPAGES fields ([aafed67](https://github.com/shadowgarden-app/bapbong/commit/aafed67))
- **layout-engine,painter-canvas:** M5 — per-page header/footer (page chrome) ([cc0cc7c](https://github.com/shadowgarden-app/bapbong/commit/cc0cc7c))
- **layout-engine,painter-canvas:** M5 — Word cell padding + underline/strike ([ff6baee](https://github.com/shadowgarden-app/bapbong/commit/ff6baee))
- **painter-canvas:** viewport page virtualization ([8c25d06](https://github.com/shadowgarden-app/bapbong/commit/8c25d06))
- **painter-canvas:** overlay canvas layer — drag/blink never re-rasterize text ([19afc87](https://github.com/shadowgarden-app/bapbong/commit/19afc87))
- **selection,input-bridge:** M4 MVP — edit on canvas (caret, selection, typing, undo) ([174d292](https://github.com/shadowgarden-app/bapbong/commit/174d292))
- **painter-canvas:** M3 — canvas viewer (ResolvedLayout → <canvas>) + playground wiring ([6a7f686](https://github.com/shadowgarden-app/bapbong/commit/6a7f686))

### 🩹 Fixes

- **painter:** table borders stroke above run shading, not under it ([51d90a6](https://github.com/shadowgarden-app/bapbong/commit/51d90a6))
- **layout,painter:** header/footer floats position in the chrome band (M14-P2b) ([b437fff](https://github.com/shadowgarden-app/bapbong/commit/b437fff))
- **painter-canvas:** one canvas per page — long docs were blank past ~57 pages ([aa10f62](https://github.com/shadowgarden-app/bapbong/commit/aa10f62))
- **painter-canvas,docx,model,layout-engine:** tables are borderless unless w:tblBorders says otherwise ([39a8ec6](https://github.com/shadowgarden-app/bapbong/commit/39a8ec6))

### 🔥 Performance

- **playground,painter-canvas:** cut per-mousemove work 53x during drag ([d2a6b9e](https://github.com/shadowgarden-app/bapbong/commit/d2a6b9e))

### 🧱 Updated Dependencies

- Updated @shadow-garden/bapbong-contracts to 0.1.0

### ❤️ Thank You

- Claude Opus 4.8 (1M context)
- Le Phuoc Minh