# @shadow-garden/bapbong-docx

DOCX (OOXML) **import and export**. Unzips a `.docx`, parses
`word/document.xml` and friends, and maps the property cascade (docDefaults →
table style → paragraph style → character style → inline) into the model
schema — then writes it back.

- **Scope:** `scope:io`
- **Depends on:** `@shadow-garden/bapbong-model`, `@shadow-garden/bapbong-contracts`, `jszip`

## What it covers

- Paragraphs, runs and marks (bold/italic/underline/strike, color, highlight,
  super/subscript), multilevel **lists** (`numbering.xml`), **tables** (borders,
  shading, vAlign, merged cells), inline **images**, **hyperlinks**.
- **Table styles**, including `w:tblStylePr` conditional formatting — first/last
  row, first/last column, row and column banding, corner cells — across all four
  channels (`w:pPr`, `w:rPr`, `w:tcPr`, cell borders), gated by `w:tblLook`.
- Page geometry from `sectPr`, **multi-column** sections, **headers/footers**
  (default / first / even), **footnotes/endnotes**, track-changes runs.
- **Comments:** `comments.xml` (range marks) + threaded/resolved metadata from
  **`commentsExtended.xml`** (w15 `paraIdParent` → `parentId`, `w15:done` → resolved).

## Public API

```ts
import { importDocx } from '@shadow-garden/bapbong-docx';

const {
  doc,
  headers,
  footers,
  footnotes,
  comments,
  page,
  titlePg,
  evenAndOdd,
} = await importDocx(fileBytes); // DocxInput = ArrayBuffer | Uint8Array | Blob
```

```ts
import { exportDocx } from '@shadow-garden/bapbong-docx';

const bytes = await exportDocx(doc, { carry: raw }); // `raw` from importDocx
```

`exportDocx` round-trips the document; passing the importer's `raw` package
back lets unmodelled parts and properties survive the save verbatim.

## Word vs the standard

Fidelity means matching **Word**, which is not always the same as matching
ISO/IEC 29500. Where the two disagree this converter follows Word and cites the
source at the call site. The ones that shaped it most:

| Behaviour                                                                                                                                       | Source                               |
| ----------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------ |
| `w:tblStylePr type="wholeTable"` is **discarded** — "Word does not apply and discards on save any properties within" it                         | [MS-OI29500] §17.18.89(a)            |
| Toggle properties (`w:b`, `w:i`, …) **do not toggle** between style layers; Word "resets the value … to the value specified"                    | [MS-OI29500] §2.1.258                |
| A missing `w:tblLook` means the bitmask **0x04A0**, not 0x0000; and `w:val` is read only when none of the six attributes is present             | [MS-OI29500] §17.4.55                |
| `tblStyleRowBandSize`/`ColBandSize` default to **0** (the standard says 1), and 0 means no banding at all                                       | [MS-OI29500] §2.1.251                |
| Conditional formats apply bands → first/last column → first/last row → corners (the standard orders it differently)                             | [MS-OI29500] §2.1.250                |
| Auto paragraph spacing is 14pt (280 twips), and only between two neighbouring paragraphs — never at a cell's edges or between items of one list | Word help; Aspose.Words; LibreOffice |

[MS-OI29500]: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/1fd4a662-8623-49c0-82f0-18fa91b413b8

The 14pt is not a guess: Aspose.Words and Word's own help both name it, and
LibreOffice's `writerfilter/DomainMapper.cxx` hardcodes `default_spacing = 280`
twips for the same setting. One MS Q&A answer describes it as 1em instead —
three sources against one, and no document in the corpus can referee it.

Row and column banding excludes the header row (and the first column) from the
count. That rule is in no document we could find; it was settled by diffing our
computation against all 268 `w:cnfStyle` strings **Word itself wrote** into a
styled 39-row table.

**Erratum.** Commit `6d2c3ec` ("a table style's `w:pPr` reaches the paragraphs
inside the table") states in its message that many real table styles keep their
`pPr` inside `tblStylePr type="wholeTable"`, and that the commit therefore does
not reach them. That is wrong: Word ignores that branch entirely (first row
above), so whole-table formatting only ever lives at the style's own
`w:pPr`/`w:rPr`/`w:tblPr` — which that commit does read. The message is left as
written; this is the correction.

## Build / test

```sh
pnpm nx build @shadow-garden/bapbong-docx
pnpm nx test  @shadow-garden/bapbong-docx
```
