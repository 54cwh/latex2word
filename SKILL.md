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
          → [mapper.js: AST → docx-js Document with __MATH_N__ placeholders]
          → [inject-omml.py: pandoc converts formulas.tex → OMML,
              inject OMML into placeholder docx, set Cambria Math font]
          → output.docx
```

## Test Files

```bash
node scripts/latex2docx.js /tmp/latex2word-test/paper.tex /tmp/paper.docx
node scripts/latex2docx.js /tmp/latex2word-test/thesis/main.tex /tmp/thesis.docx
node scripts/latex2docx.js /tmp/latex2word-test/complex.tex /tmp/complex.docx
```

## Capabilities

| Feature | How it works |
|---|---|
| Sections (\\section, \\subsection) | pandoc Header → Heading1/2/3 styles |
| Inline/Display Math | pandoc Math → OMML via inject-omml.py (pandoc → OMML post-injection). Cambria Math font set globally. |
| Tables | pandoc Table → docx-js Table with **three-line style** (top/bottom heavy lines, header bottom line, no data row borders) |
| Header/Footer | pandoc fancyhdr → docx-js Header/Footer with PAGE fields |
| Table of Contents | pandoc \\tableofcontents → docx-js TOC field code (TOC \\h \\o "1-3" \\u). SDT unwrapped for WPS compatibility. Press F9 to generate. |
| Cover page | pandoc Meta title/author/date + \\maketitle detection |
| Footnotes | pandoc Note → docx-js FootnoteReferenceRun |
| Bibliography | pandoc thebibliography RawBlock → numbered list |
| Citations | pandoc Cite → bibitem number mapping (`[1]`) |
| Theorem/Proof | pandoc Div with class → bold "Theorem."/"Proof." + content |
| Code blocks | pandoc CodeBlock → monospace paragraph |
| Figures | pandoc Figure → ImageRun with caption |
| Lists | pandoc BulletList/OrderedList → docx-js lists |
| DisplayMath | Split to separate centered paragraphs |
| Multi-file | preprocess.js merges \\input/\\include with path resolution |
| Custom commands | preprocess.js expands \\newcommand (zero-arg) |
| Bold/Italic/Math nesting | mapper.js preserves Math elements inside Strong/Emph formatting |

## Scripts

### `latex2docx.js` — Main entry

Orchestrates the full pipeline.

### `preprocess.js` — Preprocessor

- Merges \\input{file} and \\include{file} (resolves relative paths)
- Expands \\newcommand (zero-argument commands)
- Detects \\tableofcontents, \\maketitle, \\printbibliography, \\thebibliography
- Detects fancyhdr and extracts \\lhead/\\chead/\\rhead/\\lfoot/\\cfoot/\\rfoot
- Strips preamble for pandoc

### `mapper.js` — pandoc AST → docx-js Document

Handles all major pandoc block types (Header, Para, Plain, BulletList, OrderedList, Table, CodeBlock, BlockQuote, HorizontalRule, Div, Figure, RawBlock) and inline types (Str, Space, Strong, Emph, Underline, SmallCaps, Strikeout, Superscript, Subscript, Code, Link, Note, Cite, Math, RawInline, SoftBreak, LineBreak). Math elements get `__MATH_N__` placeholders for later OMML injection.

Three-line table: `tblBorders` top/bottom sz=12, header cells bottom sz=6, no data row borders.

### `inject-omml.py` — OMML post-injection

- Extracts formulas from the docx-js placeholder docx
- Runs pandoc on the collected formulas (as a single formulas.tex)
- Extracts `<m:oMath>` elements from pandoc's OMML output
- Injects them back replacing `__MATH_N__` placeholders
- Unwraps TOC SDT wrapper (for WPS compatibility)
- Sets Cambria Math as default formula font via settings.xml
