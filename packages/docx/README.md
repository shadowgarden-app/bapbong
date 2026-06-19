# @shadow-garden/bapbong-docx

DOCX (OOXML) **import** → bapbong's ProseMirror document. Unzips a `.docx`,
parses `word/document.xml` and friends, and maps the run-property cascade
(docDefaults → style → inline) into the model schema.

- **Scope:** `scope:io`
- **Depends on:** `@shadow-garden/bapbong-model`, `@shadow-garden/bapbong-contracts`, `jszip`

## What it covers

- Paragraphs, runs and marks (bold/italic/underline/strike, color, highlight,
  super/subscript), multilevel **lists** (`numbering.xml`), **tables** (borders,
  shading, vAlign, merged cells), inline **images**, **hyperlinks**.
- Page geometry from `sectPr`, **multi-column** sections, **headers/footers**
  (default / first / even), **footnotes/endnotes**, track-changes runs.
- **Comments:** `comments.xml` (range marks) + threaded/resolved metadata from
  **`commentsExtended.xml`** (w15 `paraIdParent` → `parentId`, `w15:done` → resolved).

## Public API

```ts
import { importDocx } from '@shadow-garden/bapbong-docx';

const { doc, headers, footers, footnotes, comments, page, titlePg, evenAndOdd } =
  await importDocx(fileBytes); // DocxInput = ArrayBuffer | Uint8Array | Blob
```

Export (round-trip) is a future milestone (M6).

## Build / test

```sh
pnpm nx build @shadow-garden/bapbong-docx
pnpm nx test  @shadow-garden/bapbong-docx
```
