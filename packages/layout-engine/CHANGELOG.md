## 0.1.0 (2026-07-15)

### 🚀 Features

- **layout,painter:** contain rotated images vertically, clip horizontally (M16-V2) ([491656c](https://github.com/shadowgarden-app/bapbong/commit/491656c))
- **model,layout,painter,docx:** image rotation pipeline (M16-R4) ([77573a7](https://github.com/shadowgarden-app/bapbong/commit/77573a7))
- **layout,selection:** image hit-testing foundation for resize (M16-R1) ([d5444ac](https://github.com/shadowgarden-app/bapbong/commit/d5444ac))
- **layout,docx,painter:** render text inside wps:txbx textboxes ([2c25141](https://github.com/shadowgarden-app/bapbong/commit/2c25141))
- **layout,painter:** anchored floats position inside table cells (M14-P2b) ([4bf0aff](https://github.com/shadowgarden-app/bapbong/commit/4bf0aff))
- **docx,layout,painter:** drawn shapes — wps rect/line render + round-trip (M14-P2) ([212489f](https://github.com/shadowgarden-app/bapbong/commit/212489f))
- semantic heading levels 1–6 (B3.4) ([6198747](https://github.com/shadowgarden-app/bapbong/commit/6198747))
- **docx,painter-canvas,playground:** comments — import + sidebar (read-only) ([bccb803](https://github.com/shadowgarden-app/bapbong/commit/bccb803))
- **layout-engine,docx,painter-canvas:** first/even header & footer variants ([58a51f3](https://github.com/shadowgarden-app/bapbong/commit/58a51f3))
- **layout-engine:** balance columns on final page + reserve table-cell footnotes ([dbb1880](https://github.com/shadowgarden-app/bapbong/commit/dbb1880))
- **layout-engine,docx:** multi-column sections (w:cols) ([69a0b76](https://github.com/shadowgarden-app/bapbong/commit/69a0b76))
- **layout-engine,docx,painter-canvas:** page-bottom footnote layout ([3314dc4](https://github.com/shadowgarden-app/bapbong/commit/3314dc4))
- **docx,model,layout-engine,painter-canvas:** T2 — per-cell borders + symbols ([928d44d](https://github.com/shadowgarden-app/bapbong/commit/928d44d))
- **model,docx,layout-engine:** T2 — table align, row height & cell vAlign ([731571b](https://github.com/shadowgarden-app/bapbong/commit/731571b))
- **model,docx,layout-engine,painter-canvas:** T2 — superscript / subscript ([3c227a8](https://github.com/shadowgarden-app/bapbong/commit/3c227a8))
- **model,docx,layout-engine,painter-canvas:** T1 — shading & highlight ([6149604](https://github.com/shadowgarden-app/bapbong/commit/6149604))
- **model,docx,layout-engine:** T1 — line breaks (w:br) + page breaks ([00ad335](https://github.com/shadowgarden-app/bapbong/commit/00ad335))
- **docx,model,layout-engine:** T1 — paragraph spacing (w:spacing) ([a5795fd](https://github.com/shadowgarden-app/bapbong/commit/a5795fd))
- **model,docx,layout-engine,input-bridge:** live list numbering — Enter renumbers ([8962000](https://github.com/shadowgarden-app/bapbong/commit/8962000))
- **layout-engine,docx,model:** M5 — custom tab stops (w:tabs) with leaders ([ddeefff](https://github.com/shadowgarden-app/bapbong/commit/ddeefff))
- **layout-engine,docx,model,painter-canvas:** M5 — floating images with text wrap ([bcab713](https://github.com/shadowgarden-app/bapbong/commit/bcab713))
- **docx,layout-engine,measuring:** M5 — tblCellMar, cross-run kerning, web-font loading ([44af582](https://github.com/shadowgarden-app/bapbong/commit/44af582))
- **docx,model,layout-engine,painter-canvas:** M5 — live PAGE/NUMPAGES fields ([aafed67](https://github.com/shadowgarden-app/bapbong/commit/aafed67))
- **layout-engine,painter-canvas:** M5 — per-page header/footer (page chrome) ([cc0cc7c](https://github.com/shadowgarden-app/bapbong/commit/cc0cc7c))
- **layout-engine,painter-canvas:** M5 — Word cell padding + underline/strike ([ff6baee](https://github.com/shadowgarden-app/bapbong/commit/ff6baee))
- **layout-engine,docx,model:** M5 — repeat w:tblHeader rows on every page ([cfe839c](https://github.com/shadowgarden-app/bapbong/commit/cfe839c))
- **layout-engine:** M5 — row-level table pagination + mid-row splitting ([6340e29](https://github.com/shadowgarden-app/bapbong/commit/6340e29))
- **layout-engine:** incremental re-layout via paragraph-level LayoutCache ([44612af](https://github.com/shadowgarden-app/bapbong/commit/44612af))
- **selection,input-bridge:** M4 MVP — edit on canvas (caret, selection, typing, undo) ([174d292](https://github.com/shadowgarden-app/bapbong/commit/174d292))
- **layout-engine:** M2 — inline image, table, real font metrics & tab stops ([9021f26](https://github.com/shadowgarden-app/bapbong/commit/9021f26))
- **layout-engine:** M2 — paragraph alignment, indent & hanging-indent ([abed9a9](https://github.com/shadowgarden-app/bapbong/commit/abed9a9))
- **layout-engine:** M2 — PM doc to paint-ready ResolvedLayout ([6189fcc](https://github.com/shadowgarden-app/bapbong/commit/6189fcc))

### 🩹 Fixes

- **layout,docx:** split tall table rows across pages unless w:cantSplit ([6dbd1a1](https://github.com/shadowgarden-app/bapbong/commit/6dbd1a1))
- **layout:** break words at character level when the band cannot fit them ([04bc533](https://github.com/shadowgarden-app/bapbong/commit/04bc533))
- **layout:** typed leading spaces render on first lines ([d577b7c](https://github.com/shadowgarden-app/bapbong/commit/d577b7c))
- **layout,painter:** header/footer floats position in the chrome band (M14-P2b) ([b437fff](https://github.com/shadowgarden-app/bapbong/commit/b437fff))
- **layout:** scale a too-wide table down to the column width ([13a19a2](https://github.com/shadowgarden-app/bapbong/commit/13a19a2))
- **layout-engine:** honor rowspan when assigning table columns ([17277db](https://github.com/shadowgarden-app/bapbong/commit/17277db))
- **painter-canvas,docx,model,layout-engine:** tables are borderless unless w:tblBorders says otherwise ([39a8ec6](https://github.com/shadowgarden-app/bapbong/commit/39a8ec6))
- **layout-engine:** start an oversize table row in the leftover page space ([b585053](https://github.com/shadowgarden-app/bapbong/commit/b585053))

### 🔥 Performance

- **layout-engine:** cache table layout by PM node identity ([3db8293](https://github.com/shadowgarden-app/bapbong/commit/3db8293))

### 🧱 Updated Dependencies

- Updated @shadow-garden/bapbong-contracts to 0.1.0
- Updated @shadow-garden/bapbong-model to 0.1.0

### ❤️ Thank You

- Claude Opus 4.8 (1M context)
- Le Phuoc Minh