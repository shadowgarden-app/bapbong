## 0.1.0 (2026-07-15)

### 🚀 Features

- **model,layout,painter,docx:** image rotation pipeline (M16-R4) ([77573a7](https://github.com/shadowgarden-app/bapbong/commit/77573a7))
- **painter,docx:** draw ellipse, roundRect, rightArrow, horizontalScroll ([33f4b72](https://github.com/shadowgarden-app/bapbong/commit/33f4b72))
- **docx:** flatten wpg group pictures into per-member floats ([db4bc5e](https://github.com/shadowgarden-app/bapbong/commit/db4bc5e))
- **docx:** unwrap w:sdt content controls on import ([75316be](https://github.com/shadowgarden-app/bapbong/commit/75316be))
- **layout,docx,painter:** render text inside wps:txbx textboxes ([2c25141](https://github.com/shadowgarden-app/bapbong/commit/2c25141))
- **docx:** flatten OMML equations to text runs on import ([1668740](https://github.com/shadowgarden-app/bapbong/commit/1668740))
- **docx:** import legacy VML images (w:object / w:pict + v:imagedata) ([723c7dd](https://github.com/shadowgarden-app/bapbong/commit/723c7dd))
- **docx,layout,painter:** drawn shapes — wps rect/line render + round-trip (M14-P2) ([212489f](https://github.com/shadowgarden-app/bapbong/commit/212489f))
- **docx:** rescue w:drawing wrapped in mc:AlternateContent (M14-P1) ([8f0c8cc](https://github.com/shadowgarden-app/bapbong/commit/8f0c8cc))
- semantic heading levels 1–6 (B3.4) ([6198747](https://github.com/shadowgarden-app/bapbong/commit/6198747))
- **headless:** bapbong-headless backend façade (M8) ([c15ebdc](https://github.com/shadowgarden-app/bapbong/commit/c15ebdc))
- table borders carry width/style/colour (model + import/export + painter) ([e9bc931](https://github.com/shadowgarden-app/bapbong/commit/e9bc931))
- **comments:** plugin owns the comment mark — full isolation (P4c) ([86b10b4](https://github.com/shadowgarden-app/bapbong/commit/86b10b4))
- **docx:** export section breaks + footnote references (M6 E6) ([7d0a2e6](https://github.com/shadowgarden-app/bapbong/commit/7d0a2e6))
- **docx:** re-attach body sectPr on export so headers/page round-trip ([357497c](https://github.com/shadowgarden-app/bapbong/commit/357497c))
- **docx:** carry original parts on export (M6 E4) ([6d3e915](https://github.com/shadowgarden-app/bapbong/commit/6d3e915))
- **docx:** export comments round-trip (M6 E3) ([6a6ff5b](https://github.com/shadowgarden-app/bapbong/commit/6a6ff5b))
- **docx:** export lists, tables, images & hyperlinks (M6 E2) ([c9502a6](https://github.com/shadowgarden-app/bapbong/commit/c9502a6))
- **docx:** exportDocx — paragraphs/runs/marks round-trip (M6 E1) ([4d5ed7c](https://github.com/shadowgarden-app/bapbong/commit/4d5ed7c))
- **docx:** import threaded + resolved comments (commentsExtended.xml) ([627b429](https://github.com/shadowgarden-app/bapbong/commit/627b429))
- **comments:** authoring — add/reply/resolve/delete/edit (Phase A) ([3372c94](https://github.com/shadowgarden-app/bapbong/commit/3372c94))
- **docx,painter-canvas,playground:** comments — import + sidebar (read-only) ([bccb803](https://github.com/shadowgarden-app/bapbong/commit/bccb803))
- **layout-engine,docx,painter-canvas:** first/even header & footer variants ([58a51f3](https://github.com/shadowgarden-app/bapbong/commit/58a51f3))
- **layout-engine,docx:** multi-column sections (w:cols) ([69a0b76](https://github.com/shadowgarden-app/bapbong/commit/69a0b76))
- **layout-engine,docx,painter-canvas:** page-bottom footnote layout ([3314dc4](https://github.com/shadowgarden-app/bapbong/commit/3314dc4))
- **docx:** T3 — footnotes / endnotes (numbered refs + appended notes) ([b101128](https://github.com/shadowgarden-app/bapbong/commit/b101128))
- **docx:** T3 — accept tracked changes (w:ins kept, w:del dropped) ([ff2ee6c](https://github.com/shadowgarden-app/bapbong/commit/ff2ee6c))
- **docx,model,layout-engine,painter-canvas:** T2 — per-cell borders + symbols ([928d44d](https://github.com/shadowgarden-app/bapbong/commit/928d44d))
- **model,docx,layout-engine:** T2 — table align, row height & cell vAlign ([731571b](https://github.com/shadowgarden-app/bapbong/commit/731571b))
- **model,docx,layout-engine,painter-canvas:** T2 — superscript / subscript ([3c227a8](https://github.com/shadowgarden-app/bapbong/commit/3c227a8))
- **model,docx,layout-engine,painter-canvas:** T1 — shading & highlight ([6149604](https://github.com/shadowgarden-app/bapbong/commit/6149604))
- **model,docx,layout-engine:** T1 — line breaks (w:br) + page breaks ([00ad335](https://github.com/shadowgarden-app/bapbong/commit/00ad335))
- **docx,playground:** T1 — import page size & margins from w:sectPr ([782068c](https://github.com/shadowgarden-app/bapbong/commit/782068c))
- **docx,model,layout-engine:** T1 — paragraph spacing (w:spacing) ([a5795fd](https://github.com/shadowgarden-app/bapbong/commit/a5795fd))
- **model,docx,layout-engine,input-bridge:** live list numbering — Enter renumbers ([8962000](https://github.com/shadowgarden-app/bapbong/commit/8962000))
- **layout-engine,docx,model:** M5 — custom tab stops (w:tabs) with leaders ([ddeefff](https://github.com/shadowgarden-app/bapbong/commit/ddeefff))
- **layout-engine,docx,model,painter-canvas:** M5 — floating images with text wrap ([bcab713](https://github.com/shadowgarden-app/bapbong/commit/bcab713))
- **docx,layout-engine,measuring:** M5 — tblCellMar, cross-run kerning, web-font loading ([44af582](https://github.com/shadowgarden-app/bapbong/commit/44af582))
- **docx,model,layout-engine,painter-canvas:** M5 — live PAGE/NUMPAGES fields ([aafed67](https://github.com/shadowgarden-app/bapbong/commit/aafed67))
- **docx:** M5 — cascade pPr (w:jc/w:ind/w:numPr) from paragraph styles ([3adc2a2](https://github.com/shadowgarden-app/bapbong/commit/3adc2a2))
- **layout-engine,docx,model:** M5 — repeat w:tblHeader rows on every page ([cfe839c](https://github.com/shadowgarden-app/bapbong/commit/cfe839c))
- **layout-engine:** M2 — paragraph alignment, indent & hanging-indent ([abed9a9](https://github.com/shadowgarden-app/bapbong/commit/abed9a9))
- **docx:** M1+ — header/footer parts ([13c663b](https://github.com/shadowgarden-app/bapbong/commit/13c663b))
- **docx:** M1+ — theme color resolution (w:themeColor) ([029ee0e](https://github.com/shadowgarden-app/bapbong/commit/029ee0e))
- **docx,model:** M1+ — hyperlinks + inline images ([9fa6c02](https://github.com/shadowgarden-app/bapbong/commit/9fa6c02))
- **docx:** M1+ — collapse vertical cell merges (w:vMerge) into rowspan ([e3d1b88](https://github.com/shadowgarden-app/bapbong/commit/e3d1b88))
- **docx,model:** M1+ — tables + document-order traversal ([8362cf0](https://github.com/shadowgarden-app/bapbong/commit/8362cf0))
- **docx,model:** M1+ — lists / multilevel numbering ([b7f6607](https://github.com/shadowgarden-app/bapbong/commit/b7f6607))
- **docx,model:** M1+ — run-property cascade + color/size/font ([e173d46](https://github.com/shadowgarden-app/bapbong/commit/e173d46))
- **docx,model:** M1 — DOCX import to ProseMirror model ([07d236b](https://github.com/shadowgarden-app/bapbong/commit/07d236b))

### 🩹 Fixes

- **layout,docx:** split tall table rows across pages unless w:cantSplit ([6dbd1a1](https://github.com/shadowgarden-app/bapbong/commit/6dbd1a1))
- **docx:** stop strnum from mangling numeric-looking document text ([0743e46](https://github.com/shadowgarden-app/bapbong/commit/0743e46))
- **docx:** per-level list indents from the numbering definition (w:lvl/w:pPr) ([1f9dc97](https://github.com/shadowgarden-app/bapbong/commit/1f9dc97))
- **painter-canvas,docx,model,layout-engine:** tables are borderless unless w:tblBorders says otherwise ([39a8ec6](https://github.com/shadowgarden-app/bapbong/commit/39a8ec6))

### 🧱 Updated Dependencies

- Updated @shadow-garden/bapbong-contracts to 0.1.0
- Updated @shadow-garden/bapbong-model to 0.1.0

### ❤️ Thank You

- Claude Opus 4.8 (1M context)
- Le Phuoc Minh