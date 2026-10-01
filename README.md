# latex2word

Convert LaTeX documents to Word `.docx` files that still look like papers.

`pandoc input.tex -o output.docx` gives you a document, but not a paper: tables come out boxed, equations aren't numbered, theorem and proof environments are plain text, and the title page and running header are missing.

latex2word keeps pandoc for parsing and for rendering formulas, but builds the Word file itself with [docx-js](https://github.com/dolanmiu/docx). That gives it control over the things pandoc's docx writer won't do: three-line tables, `sec-seq` numbering, theorem/proof styling, a real TOC field, a cover page, and page-number headers and footers.

## What it keeps

- **Equations** — inline and display math become native Word equations (OMML) in Cambria Math, editable in Word. Numbered equations keep a `(sec-seq)` number.
- **Tables** — three-line (booktabs) style: heavy rules top and bottom, one rule under the header, nothing between data rows. Captions are auto-numbered.
- **Figures** — embedded with auto-numbered captions.
- **Headings** — `\section` / `\subsection` map to Word heading styles. Unnumbered sections such as Abstract stay out of the number sequence.
- **Citations and bibliography** — `\cite` maps to the matching `\bibitem` number; the bibliography becomes a numbered list.
- **Footnotes, TOC, cover page, running header/footer, theorem/proof environments, code blocks, lists.**
- **Multi-file projects** — `\input` and `\include` are merged with relative paths resolved.
- **Zero-argument `\newcommand` macros** are expanded before conversion.

## Requirements

- Node.js 18+
- [pandoc](https://pandoc.org/) on your `PATH`
- Python 3 (for `scripts/inject-omml.py`)

## Install

```bash
git clone https://github.com/54cwh/latex2word.git
cd latex2word
npm install
```

## Use

```bash
node scripts/latex2docx.js thesis.tex thesis.docx
```

Omit the output path and it writes `thesis.docx` next to the input.

## How it works

```
input.tex
  → scripts/preprocess.js    merge \input/\include, expand \newcommand,
                             detect TOC/title/bibliography, extract fancyhdr
  → pandoc -f latex+raw_tex -t json
  → scripts/latexmlibs/      walk the pandoc AST, build a docx-js Document,
                             leave __MATH_N__ placeholders for formulas
  → scripts/inject-omml.py   render the formulas with pandoc, splice the OMML
                             back into the placeholders, set Cambria Math
  → output.docx
```

## Layout

```
scripts/
  latex2docx.js      main entry: runs the whole pipeline
  preprocess.js      LaTeX preprocessing
  inject-omml.py     OMML post-processing
  latexmlibs/        pandoc AST → docx-js modules
    index.js         AST walker
    blocks.js        headers, tables, figures, lists, divs, code blocks
    inline.js        text runs, math, links, notes, cites, formatting
    document.js      builds the Document (styles, sections, headers, footers)
    formulas.js      formula collection and numbering
    bib.js           citations and bibliography
    state.js         shared counters and maps
    utils.js         image embedding, text extraction, style helpers
```

## Notes and limitations

- The table of contents is inserted as a Word field. Open the file and press **F9** to populate it.
- `\newcommand` expansion handles zero-argument macros only.
- The mapper covers the constructs listed above; a package or environment it doesn't know will need handling added in `latexmlibs/`.
- Formulas pass through pandoc twice, once for structure and once for OMML, so a construct pandoc can't parse won't survive the trip.

## Using it as an agent skill

This repo is also an agent skill. `SKILL.md` describes when to trigger it and what each script does. Point your agent's skills directory at this folder, or symlink it in.

## License

See [LICENSE.txt](LICENSE.txt).
