---
name: latex2word
description: "Use this skill whenever the user wants to convert LaTeX (.tex) documents into Word (.docx) files. Triggers include: any mention of 'convert latex to word', 'tex to docx', 'latex to docx', 'compile latex to word', 'turn my paper into a word document', or any request to convert academic papers, theses, math-heavy documents, or any .tex file into a polished .docx. This skill uses a direct LaTeX-to-docx pipeline (preprocess → pandoc AST → docx-js mapper → OMML injection) for high-fidelity output that preserves sections, math formulas (OMML), three-line tables, figures, footnotes, citations, TOC, bibliography, header/footer, theorem/proof environments, and multi-file projects."
license: Proprietary. LICENSE.txt has complete terms
---

# latex2word Conversion Scripts

This skill provides conversion scripts in `scripts/`. Use them instead of writing conversion code from scratch.

## Quick Start

```bash
# Convert a single .tex file
node scripts/latex2docx.js input.tex output.docx
```

If no output path is given, writes to `input.docx`.

## Pipeline

```
input.tex → [preprocess.js: merge \input/\include, expand \newcommand,
              detect TOC/title/bib, extract fancyhdr]
          → [pandoc -f latex+raw_tex -t json]
          → [scripts/latexmlibs/: modular AST → docx-js Document
             with __MATH_N__ placeholders]
          → [inject-omml.py: pandoc converts formulas.tex → OMML,
              inject OMML into placeholder docx, set Cambria Math font]
          → output.docx
```

## Capabilities

| Feature | How it works |
|---|---|
| Sections (\section, \subsection) | pandoc Header → Heading1/2/3 styles. Unnumbered sections (Abstract) have `unnumbered` class, excluded from section counter. |
| Inline/Display Math | pandoc Math → OMML via inject-omml.py (pandoc → OMML post-injection). Cambria Math font set globally. Numbered equations use `#(sec-seq)` format. |
| Tables | pandoc Table → docx-js Table with **three-line style** (top/bottom heavy lines, header bottom line, no data row borders). Auto-numbered as `Table sec-seq` in caption. |
| Header/Footer | pandoc fancyhdr → docx-js Header/Footer with PAGE fields |
| Table of Contents | pandoc \tableofcontents → docx-js TOC field code (TOC \h \o "1-3" \u). SDT unwrapped for WPS compatibility. Press F9 to generate. |
| Cover page | pandoc Meta title/author/date + \maketitle detection |
| Footnotes | pandoc Note → docx-js FootnoteReferenceRun |
| Bibliography | pandoc thebibliography RawBlock → numbered list |
| Citations | pandoc Cite → bibitem number mapping (`[1]`) |
| Theorem/Proof | pandoc Div with class → bold "Theorem."/"Proof." + content |
| Code blocks | pandoc CodeBlock → monospace paragraph |
| Figures | pandoc Figure → ImageRun with auto-numbered caption `Figure sec-seq` |
| Lists | pandoc BulletList/OrderedList → docx-js lists |
| DisplayMath | Split to separate centered paragraphs |
| Multi-file | preprocess.js merges \input/\include with path resolution |
| Custom commands | preprocess.js expands \newcommand (zero-arg) |
| Bold/Italic/Math nesting | inline.js preserves Math elements inside Strong/Emph formatting |
| Theorem/Proof/QED | pandoc Div → structured paragraph with QED symbol (∎) |
| Section auto-numbering | All numbered elements (equations, tables, figures) use `sec-seq` format, resetting per numbered Heading 1. |

## Scripts

### `latex2docx.js` — Main entry

Orchestrates the full pipeline: preprocess → pandoc AST → docx-js → OMML injection.

### `preprocess.js` — Preprocessor

- Merges \input{file} and \include{file} (resolves relative paths)
- Expands \newcommand (zero-argument commands)
- Detects \tableofcontents, \maketitle, \printbibliography, \thebibliography
- Detects fancyhdr and extracts \lhead/\chead/\rhead/\lfoot/\cfoot/\rfoot
- Strips preamble, converts abstract to \section*{Abstract}
- Strips \documentclass from body

### `scripts/latexmlibs/` — Modular pandoc AST → docx-js

| Module | Role |
|---|---|
| `index.js` | Entry point: walks pandoc AST blocks, delegates to specialized modules |
| `blocks.js` | Block-level mappers: Header, Table, Figure, Para, Lists, Div, RawBlock, CodeBlock, HorizontalRule, BlockQuote |
| `inline.js` | Inline mappers: Str, Space, Strong, Emph, Math, Link, Note, Cite, Code, formatting spans |
| `document.js` | docx-js Document builder: creates Document with styles, sections, headers, footers, cover page |
| `formulas.js` | Formula collection/numbering: extracts numbered display math, builds OMML placeholder mapping |
| `bib.js` | Bibliography/citation handling: maps bibitem citations to numbers, builds numbered reference list |
| `state.js` | Shared state: formula tracker, footnote counter, citation map, equation/section/figure/table counters |
| `utils.js` | Utilities: image embedding (detects PNG/PDF/JPG by magic bytes), text extraction, style helpers |

Three-line table style: `tblBorders` top/bottom sz=12, header cells bottom sz=6, no data row borders.

### `inject-omml.py` — OMML post-injection

- Extracts formulas from the docx-js placeholder docx
- Runs pandoc on the collected formulas (as a single formulas.tex)
- Extracts `<m:oMath>` elements from pandoc's OMML output
- Injects them back replacing `__MATH_N__` placeholders
- Unwraps TOC SDT wrapper (for WPS compatibility)
- Sets Cambria Math as default formula font via settings.xml
