## 0.1.0 (2026-07-15)

### 🚀 Features

- **model,layout,painter,docx:** image rotation pipeline (M16-R4) ([77573a7](https://github.com/shadowgarden-app/bapbong/commit/77573a7))
- **layout,docx,painter:** render text inside wps:txbx textboxes ([2c25141](https://github.com/shadowgarden-app/bapbong/commit/2c25141))
- **docx,layout,painter:** drawn shapes — wps rect/line render + round-trip (M14-P2) ([212489f](https://github.com/shadowgarden-app/bapbong/commit/212489f))
- semantic heading levels 1–6 (B3.4) ([6198747](https://github.com/shadowgarden-app/bapbong/commit/6198747))
- **comments:** plugin owns the comment mark — full isolation (P4c) ([86b10b4](https://github.com/shadowgarden-app/bapbong/commit/86b10b4))
- **comments:** @mention support in the comment composer (Phase B) ([7d7db09](https://github.com/shadowgarden-app/bapbong/commit/7d7db09))
- **comments:** authoring — add/reply/resolve/delete/edit (Phase A) ([3372c94](https://github.com/shadowgarden-app/bapbong/commit/3372c94))
- **docx,painter-canvas,playground:** comments — import + sidebar (read-only) ([bccb803](https://github.com/shadowgarden-app/bapbong/commit/bccb803))
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
- **layout-engine,docx,model:** M5 — repeat w:tblHeader rows on every page ([cfe839c](https://github.com/shadowgarden-app/bapbong/commit/cfe839c))
- **layout-engine:** M2 — paragraph alignment, indent & hanging-indent ([abed9a9](https://github.com/shadowgarden-app/bapbong/commit/abed9a9))
- **playground:** Angular demo app for importDocx ([2707b9b](https://github.com/shadowgarden-app/bapbong/commit/2707b9b))
- **docx,model:** M1+ — hyperlinks + inline images ([9fa6c02](https://github.com/shadowgarden-app/bapbong/commit/9fa6c02))
- **docx,model:** M1+ — tables + document-order traversal ([8362cf0](https://github.com/shadowgarden-app/bapbong/commit/8362cf0))
- **docx,model:** M1+ — lists / multilevel numbering ([b7f6607](https://github.com/shadowgarden-app/bapbong/commit/b7f6607))
- **docx,model:** M1+ — run-property cascade + color/size/font ([e173d46](https://github.com/shadowgarden-app/bapbong/commit/e173d46))
- **docx,model:** M1 — DOCX import to ProseMirror model ([07d236b](https://github.com/shadowgarden-app/bapbong/commit/07d236b))

### 🩹 Fixes

- **layout,docx:** split tall table rows across pages unless w:cantSplit ([6dbd1a1](https://github.com/shadowgarden-app/bapbong/commit/6dbd1a1))
- **painter-canvas,docx,model,layout-engine:** tables are borderless unless w:tblBorders says otherwise ([39a8ec6](https://github.com/shadowgarden-app/bapbong/commit/39a8ec6))

### ❤️ Thank You

- Claude Opus 4.8 (1M context)
- Le Phuoc Minh