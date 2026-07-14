const {
  Document, Paragraph, TextRun, Header, Footer, PageNumber,
  HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, HeightRule, TableLayoutType, VerticalAlign,
  ShadingType, BorderStyle, PageBreak, PageReference,
  FootnoteReferenceRun, ImageRun, ExternalHyperlink,
  InternalHyperlink, Bookmark, TableOfContents,
  LevelFormat, UnderlineType, convertInchesToTwip,
  Tab, TabStopType, SimpleField
} = require('docx');

const { Math: DocxMath, MathRun } = require('docx');
const fs = require('fs');
const path = require('path');

let formulaTracker = null;
let footnoteCounter = 0;
let footnoteTexts = {};
let citationMap = {};
let eqCounter = 0;
let sectionNumber = 1;
let imageDir = null;

function setFormulaTracker(tracker) {
  formulaTracker = tracker;
}

function expandFormatted(items, fmt) {
  return items.flatMap(item => {
    const t = item.t, c = item.c;
    switch (t) {
      case 'Str': return [new TextRun({ ...fmt, text: c })];
      case 'Space': return [new TextRun({ ...fmt, text: ' ' })];
      case 'SoftBreak': return [new TextRun({ ...fmt, text: ' ' })];
      case 'LineBreak': return [new TextRun({ ...fmt, break: 1 })];
      case 'Emph': return expandFormatted(c, { ...fmt, italics: true });
      case 'Strong': return expandFormatted(c, { ...fmt, bold: true });
      case 'Underline': return expandFormatted(c, { ...fmt, underline: { type: UnderlineType.SINGLE } });
      case 'SmallCaps': return expandFormatted(c, { ...fmt, smallCaps: true });
      case 'Strikeout': return expandFormatted(c, { ...fmt, strike: true });
      case 'Superscript': return expandFormatted(c, { ...fmt, superScript: true });
      case 'Subscript': return expandFormatted(c, { ...fmt, subScript: true });
      case 'Span': return expandFormatted(c[1] || c, fmt);
      case 'Quoted': return expandFormatted(c[1] || c, fmt);
      default: return mapInline(item);
    }
  });
}

function mapInline(inline) {
  if (!inline) return [];
  if (Array.isArray(inline)) {
    return inline.flatMap(i => mapInline(i));
  }

  const t = inline.t;
  const c = inline.c;

  switch (t) {
    case 'Str':
      return [new TextRun(c)];

    case 'Space':
      return [new TextRun(' ')];

    case 'SoftBreak':
      return [new TextRun(' ')];

    case 'LineBreak':
      return [];

    case 'Strong':
      return expandFormatted(c, { bold: true });

    case 'Emph':
      return expandFormatted(c, { italics: true });

    case 'Underline':
      return expandFormatted(c, { underline: { type: UnderlineType.SINGLE } });

    case 'SmallCaps':
      return expandFormatted(c, { smallCaps: true });

    case 'Strikeout':
      return expandFormatted(c, { strike: true });

    case 'Superscript':
      return expandFormatted(c, { superScript: true });

    case 'Subscript':
      return expandFormatted(c, { subScript: true });

    case 'Code':
      return [new TextRun({ text: c[1], font: 'Courier New' })];

    case 'Math': {
      const mathType = c[0].t || c[0];
      const isDisplay = mathType === 'DisplayMath';
      const latex = c[1];
      if (formulaTracker) {
        const idx = formulaTracker.counter++;
        formulaTracker.formulas.push({ index: idx, latex, display: isDisplay });
        return [new DocxMath({ children: [new MathRun(`__MATH_${idx}__`)] })];
      }
      return [new DocxMath({ children: [new MathRun(latex)] })];
    }

    case 'Link': {
      const url = c[2][0];
      return [new ExternalHyperlink({
        children: [new TextRun({ text: extractText(c[1]), style: 'Hyperlink' })],
        link: url
      })];
    }

    case 'Image': {
      const src = c[2][0];
      const alt = extractText(c[1]);
      const img = embedImage(src, { alt });
      if (img) return [img];
      return [new TextRun(`[Image: ${alt} — ${src}]`)];
    }

    case 'Note': {
      const noteId = ++footnoteCounter;
      const noteText = extractText(c[0] && c[0].c);
      footnoteTexts[noteId] = { children: [new Paragraph({ children: [new TextRun(noteText)] })] };
      return [new FootnoteReferenceRun(noteId)];
    }

    case 'Cite': {
      const citationId = (c[0] && c[0][0] && c[0][0].citationId) || '';
      const num = citationMap[citationId];
      if (num) return [new TextRun({ text: `[${num}]`, superScript: true })];
      return [new TextRun({ text: `[${citationId}]`, superScript: true })];
    }

    case 'Span': {
      return mapInlines(c[1]);
    }

    case 'Quoted':
      return mapInlines(c[1]);

    case 'RawInline': {
      if (c[0] === 'latex') {
        const raw = c[1];
        const cleaned = raw.replace(/\\(label|ref|cite|pageref|index|glossary)\{[^}]*\}/g, '').trim();
        if (cleaned) return [new TextRun(cleaned)];
        return [];
      }
      return [];
    }

    default:
      if (c && typeof c === 'object') {
        if (Array.isArray(c)) return [new TextRun(extractText(c))];
        return [new TextRun(JSON.stringify(c))];
      }
      return [new TextRun(String(c || ''))];
  }
}

function extractText(inlines) {
  if (!inlines) return '';
  if (typeof inlines === 'string') return inlines;
  if (Array.isArray(inlines)) {
    return inlines.map(i => extractText(i)).join('');
  }
  if (typeof inlines === 'object') {
    if (inlines.t === 'Str') return inlines.c || '';
    if (inlines.t === 'Space' || inlines.t === 'SoftBreak') return ' ';
    if (inlines.t === 'Code' || inlines.t === 'Math') return inlines.c[1] || inlines.c || '';
    if (inlines.t === 'LineBreak') return ' ';
    if (inlines.t === 'RawInline') {
      if (inlines.c[0] === 'latex') {
        return inlines.c[1].replace(/\\(label|ref|cite|pageref|index|glossary)\{[^}]*\}/g, '').trim();
      }
      return '';
    }
    if (inlines.c) return extractText(inlines.c);
  }
  return '';
}

function mapInlines(inlines) {
  if (!inlines) return [];
  return inlines.flatMap(inline => mapInline(inline));
}

function buildParagraphs(inlines) {
  if (!inlines || !inlines.length) return [new Paragraph({ children: [] })];

  const segments = [];
  let current = [];
  for (const item of inlines) {
    if (item.t === 'Math' && (item.c[0].t || item.c[0]) === 'DisplayMath') {
      if (current.length) {
        segments.push({ items: current });
        current = [];
      }
      segments.push({ displayMath: item });
    } else {
      current.push(item);
    }
  }
  if (current.length) segments.push({ items: current });

  return segments.map(seg => {
    if (seg.displayMath) {
      const latex = seg.displayMath.c[1] || '';
      const isNumbered = /\\label/.test(latex);
      if (isNumbered) {
        const eqNum = ++eqCounter;
        const latex = seg.displayMath.c[1] || '';
        return new Paragraph({
          tabStops: [
            { type: TabStopType.CENTER, position: 4513 },
            { type: TabStopType.RIGHT, position: 9026 },
          ],
          spacing: { before: 120, after: 120 },
          children: [
            new Tab(),
            buildNumberedFormula(latex, sectionNumber, eqNum),
            new Tab(),
          ]
        });
      }
      return new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120 },
        children: mapInline(seg.displayMath)
      });
    }
    return new Paragraph({ children: mapInlines(seg.items) });
  });
}

function mapAlign(align) {
  switch (align) {
    case 'AlignLeft': return AlignmentType.LEFT;
    case 'AlignCenter': return AlignmentType.CENTER;
    case 'AlignRight': return AlignmentType.RIGHT;
    case 'AlignDefault': return AlignmentType.LEFT;
    default: return AlignmentType.LEFT;
  }
}

function parseBibText(latex) {
  latex = latex.replace(/\\end\{thebibliography\}.*$/, '').trim();
  const runs = [];
  let buf = '';

  function flush() {
    if (buf) { runs.push(new TextRun({ text: buf })); buf = ''; }
  }

  function parseArgs(i) {
    const args = [];
    let depth = 0;
    while (i < latex.length && latex[i] !== '{') i++;
    if (latex[i] === '{') { depth = 1; i++; }
    let arg = '';
    while (i < latex.length && depth > 0) {
      if (latex[i] === '{') { depth++; arg += '{'; i++; }
      else if (latex[i] === '}') { depth--; if (depth > 0) arg += '}'; i++; }
      else { arg += latex[i]; i++; }
    }
    if (arg) args.push(arg);
    // Check for optional [opt]
    while (latex[i] === '[') {
      i++;
      let opt = '';
      while (i < latex.length && latex[i] !== ']') { opt += latex[i]; i++; }
      if (i < latex.length) i++;
      args.push(opt);
    }
    return { args, pos: i };
  }

  for (let i = 0; i < latex.length; i++) {
    const ch = latex[i];

    if (ch === '\\') {
      const next = latex[i + 1];
      if (!next) break;

      if (/[a-zA-Z]/.test(next)) {
        const m = latex.slice(i).match(/^\\([a-zA-Z]+)/);
        if (m) {
          const cmd = m[1];
          const cmdLen = m[0].length;
          i += cmdLen;

          const { args, pos } = parseArgs(i);
          i = pos - 1;

          flush();
          const content = args[0] || '';

          switch (cmd) {
            case 'textit': case 'emph':
              runs.push(new TextRun({ text: content, italics: true })); break;
            case 'textbf':
              runs.push(new TextRun({ text: content, bold: true })); break;
            case 'texttt':
              runs.push(new TextRun({ text: content, font: 'Courier New' })); break;
            case 'textsc':
              runs.push(new TextRun({ text: content, smallCaps: true })); break;
            case 'textcolor': {
              const color = args[0] || '';
              const text = args[1] || args[0] || '';
              if (color && args.length > 1) {
                const colorCodes = { red: 'FF0000', blue: '0000FF', green: '00AA00', gray: '888888' };
                runs.push(new TextRun({ text, color: colorCodes[color] || color }));
              } else {
                runs.push(new TextRun({ text }));
              }
              break;
            }
            case 'url':
              runs.push(new TextRun({ text: content, color: '0563C1', underline: { type: UnderlineType.SINGLE } }));
              break;
            default:
              if (content) runs.push(new TextRun({ text: content }));
              break;
          }
          continue;
        }
      }

      if (next === '{' || next === '}') {
        i += 2; continue;
      }

      if (next === '~') { buf += '\u00A0'; i += 2; continue; }
      if (next === '\\') { buf += '\\'; i += 2; continue; }
      if (next === ' ') { buf += ' '; i += 2; continue; }

      // Accent commands: \`{e}, \'{e}, \^{e}, \"{e}, \~{n}, \={e}, etc.
      if (['`', '\'', '^', '"', '~', '=', '.'].includes(next)) {
        const accentMatch = latex.slice(i).match(/^\\([`'\^"~=.])\{([^}]*)\}/);
        if (accentMatch) {
          buf += accentMatch[2];
          i += accentMatch[0].length - 1;
          continue;
        }
      }

      // Single char control sequence like \&, \%, \$, \#
      if (/[&%$#_]/.test(next)) {
        buf += next; i += 2; continue;
      }

      i++;
      continue;
    }

    if (ch === '{' || ch === '}') continue;

    if (ch === '`' && latex[i + 1] === '`') { flush(); buf += '\u201C'; i++; continue; }
    if (ch === '\'' && latex[i + 1] === '\'') { flush(); buf += '\u201D'; i++; continue; }
    if (ch === '~') { buf += '\u00A0'; continue; }

    buf += ch;
  }

  flush();
  return runs;
}

function mapBlock(block) {
  if (!block) return [];
  const t = block.t;
  const c = block.c;

  switch (t) {
    case 'Header': {
      const level = c[0];
      if (level === 1) { sectionNumber++; eqCounter = 0; }
      const inlines = c[2];
      const headingKey = `HEADING_${level}`;
      const heading = HeadingLevel[headingKey];
      return [new Paragraph({
        heading,
        children: mapInlines(inlines)
      })];
    }

    case 'Para': {
      return buildParagraphs(c);
    }

    case 'Plain': {
      return buildParagraphs(c);
    }

    case 'BulletList': {
      const items = [];
      c.forEach(itemBlocks => {
        const firstBlock = itemBlocks[0];
        if (firstBlock && firstBlock.t === 'Para') {
          items.push(new Paragraph({
            numbering: { reference: 'bullets', level: 0 },
            children: mapInlines(firstBlock.c)
          }));
        }
      });
      return items;
    }

    case 'OrderedList': {
      const items = [];
      const itemBlocks = c[1];
      itemBlocks.forEach(itemBlocksInner => {
        const firstBlock = itemBlocksInner[0];
        if (firstBlock && firstBlock.t === 'Para') {
          items.push(new Paragraph({
            numbering: { reference: 'numbers', level: 0 },
            children: mapInlines(firstBlock.c)
          }));
        }
      });
      return items;
    }

    case 'Table': {
      const tableAttr = c[0];
      const caption = c[1];
      const colspecs = c[2] || [];
      const thead = c[3] || [];
      const tbodies = c[4] || [];

      const cols = colspecs.length || 1;
      const colWidth = Math.floor(9026 / cols);

      function cellFromPandoc(cell, isHeader, isLastRow) {
        const cellAttr = cell[0];
        const cellAlign = cell[1];
        const rowspan = cell[2] || 1;
        const colspan = cell[3] || 1;
        const cellBlocks = cell[4] || [];
        const cellChildren = cellBlocks.flatMap(b => mapBlock(b));
        const opts = {
          columnSpan: colspan,
          rowSpan: rowspan,
          width: { size: colWidth * colspan, type: WidthType.DXA },
          children: cellChildren.length ? cellChildren : [new Paragraph({ children: [] })]
        };
        if (isHeader) {
          opts.borders = {
            bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000' }
          };
        }
        return new TableCell(opts);
      }

      function rowFromPandoc(rowData, isHeader) {
        const rowCells = rowData[1] || [];
        return new TableRow({
          children: rowCells.map(cell => cellFromPandoc(cell, isHeader))
        });
      }

      const rows = [];

      if (thead.length >= 2 && Array.isArray(thead[1])) {
        const headRows = thead[1];
        headRows.forEach(row => {
          rows.push(rowFromPandoc(row, true));
        });
      }

      tbodies.forEach(tbody => {
        if (!Array.isArray(tbody)) return;
        const bodyRows = Array.isArray(tbody[3]) ? tbody[3] : [];
        bodyRows.forEach(row => {
          if (Array.isArray(row) && row.length >= 2 && Array.isArray(row[1])) {
            rows.push(rowFromPandoc(row));
          }
        });
        if (Array.isArray(tbody[2])) {
          tbody[2].forEach(row => {
            if (Array.isArray(row) && row.length >= 2 && Array.isArray(row[1])) {
              rows.push(rowFromPandoc(row, false));
            }
          });
        }
      });

      const result = [];
      if (caption && caption[1]) {
        const captionText = extractText(caption[1]);
        if (captionText) {
          result.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 120 },
            children: [new TextRun({ text: captionText, bold: true, italics: true, size: 20 })]
          }));
        }
      }
      result.push(new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: Array(cols).fill(colWidth),
        borders: {
          top: { style: BorderStyle.SINGLE, size: 12, color: '000000' },
          bottom: { style: BorderStyle.SINGLE, size: 12, color: '000000' },
          insideHorizontal: { style: BorderStyle.NONE },
          insideVertical: { style: BorderStyle.NONE },
          left: { style: BorderStyle.NONE },
          right: { style: BorderStyle.NONE }
        },
        rows
      }));
      result.push(new Paragraph({ children: [new TextRun('')] }));
      return result;
    }

    case 'CodeBlock': {
      const codeText = c[1];
      const lang = (c[0] && c[0][1] && c[0][1][0]) || '';
      const result = [];
      if (lang) {
        result.push(new Paragraph({
          spacing: { before: 120 },
          children: [new TextRun({ text: `[${lang}]`, size: 18, color: '666666', italics: true })]
        }));
      }
      codeText.split('\n').forEach(line => {
        result.push(new Paragraph({
          spacing: { before: 0, after: 0 },
          shading: { type: ShadingType.CLEAR, fill: 'F0F0F0' },
          indent: { left: 360 },
          children: [new TextRun({ text: line, font: 'Courier New', size: 18 })]
        }));
      });
      result.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
      return result;
    }

    case 'BlockQuote': {
      return c.flatMap(b => {
        const blocks = mapBlock(b);
        return blocks.map(p => {
          if (p instanceof Paragraph) {
            return new Paragraph({
              indent: { left: 720, right: 720 },
              spacing: { before: 80, after: 80 },
              children: p.children
            });
          }
          return p;
        });
      });
    }

    case 'HorizontalRule': {
      return [new Paragraph({
        border: {
          bottom: { style: BorderStyle.SINGLE, size: 6, color: '999999', space: 1 }
        },
        spacing: { before: 200, after: 200 }
      })];
    }

    case 'Div': {
      const classes = c[0][1] || [];
      const divBlocks = c[1];

      if (classes.includes('theorem')) {
        const firstInline = extractText(divBlocks[0] && divBlocks[0].c);
        const contentBlocks = divBlocks.flatMap(b => mapBlock(b));
        return [
          new Paragraph({
            spacing: { before: 120, after: 60 },
            children: [new TextRun({ text: 'Theorem.', bold: true, italics: true })].concat(contentBlocks[0] ? contentBlocks[0].children || [] : [])
          }),
          ...contentBlocks.slice(1)
        ];
      }

      if (classes.includes('proof')) {
        const contentBlocks = divBlocks.flatMap(b => mapBlock(b));
        if (contentBlocks.length === 0) return [];
        const first = contentBlocks[0];
        const firstChildren = first && first.children ? first.children : [];
        const proofFirst = new Paragraph({
          spacing: { before: 120, after: 60 },
          children: [new TextRun({ text: 'Proof.', italics: true }),
                     new TextRun(' '), ...firstChildren]
        });
        const rest = contentBlocks.slice(1).map((b, i, arr) => {
          if (i === arr.length - 1 && b.children) {
            return new Paragraph({
              spacing: { before: 60, after: 60 },
              children: [...b.children, new TextRun(' □')]
            });
          }
          return b;
        });
        return [proofFirst, ...rest];
      }

      return divBlocks.flatMap(b => mapBlock(b));
    }

    case 'RawBlock': {
      if (c[0] === 'latex' && c[1].includes('\\begin{thebibliography}')) {
        const bibItems = c[1].match(/\\bibitem\{[^}]*\}\s*([^]*?)(?=\\bibitem|\\end\{thebibliography\}|$)/g) || [];
        const items = bibItems.map((item, i) => {
          const text = item.replace(/\\bibitem\{[^}]*\}\s*/, '').trim();
          return new Paragraph({
            spacing: { before: 60, after: 60 },
            children: [new TextRun({ text: `[${i + 1}] `, bold: true }),
                       ...parseBibText(text)]
          });
        });
        return [
          new Paragraph({
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun('Bibliography')]
          }),
          ...items
        ];
      }
      return [];
    }

    case 'Figure': {
      const figBlocks = c[2];
      const caption = c[1];
      const capText = extractText(caption);
      const result = [];
      figBlocks.forEach(b => {
        if (b.t === 'Image') {
          const src = b.c[2][0];
          const alt = extractText(b.c[1]);
          const img = embedImage(src, { alt, width: 500 });
          if (img) {
            result.push(new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [img]
            }));
          } else {
            result.push(new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun(`[Figure: ${capText}] (${src})`)]
            }));
          }
        } else {
          result.push(...mapBlock(b));
        }
      });
      if (capText) {
        result.push(new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 60 },
          children: [new TextRun({ text: capText, italics: true, size: 20 })]
        }));
      }
      return result;
    }

    case 'LineBlock':
      return [new Paragraph({ children: mapInlines(c) })];

    default:
      return [];
  }
}

function mapBlocks(blocks) {
  return blocks.flatMap(block => mapBlock(block));
}

function collectFormulasFromAst(blocks) {
  const formulas = [];
  function walkInlines(items) {
    if (!items) return;
    if (Array.isArray(items)) { items.forEach(walkInlines); return; }
    if (typeof items !== 'object') return;
    if (items.t === 'Math') {
      const mt = items.c[0].t || items.c[0];
      formulas.push({ latex: items.c[1], display: mt === 'DisplayMath' });
    }
    if (items.c) {
      if (Array.isArray(items.c)) items.c.forEach(walkInlines);
      else walkInlines(items.c);
    }
  }
  function walkBlocks(list) {
    if (!Array.isArray(list)) return;
    list.forEach(b => {
      if (!b) return;
      switch (b.t) {
        case 'Para': case 'Plain': walkInlines(b.c); break;
        case 'Header': walkInlines(b.c[2]); break;
        case 'BulletList': b.c.forEach(i => walkBlocks(i)); break;
        case 'OrderedList': b.c[1].forEach(i => walkBlocks(i)); break;
        case 'Div': case 'BlockQuote': walkBlocks(b.c[1] || b.c); break;
        case 'Figure': walkBlocks(b.c[2]); break;
        case 'Table': {
          [b.c[3], b.c[4]].forEach(part => {
            if (!Array.isArray(part)) return;
            part.forEach(row => {
              if (Array.isArray(row)) {
                const cells = Array.isArray(row[1]) ? row[1] : row;
                cells.forEach(cell => {
                  if (cell && cell[4]) walkBlocks(cell[4]);
                });
              }
            });
          });
          break;
        }
      }
    });
  }
  walkBlocks(Array.isArray(blocks) ? blocks : (blocks.blocks || []));
  return formulas;
}

function formulasToTex(formulas) {
  const lines = [
    '\\documentclass{article}',
    '\\usepackage{amsmath,amssymb}',
    '\\begin{document}',
  ];
  for (const f of formulas) {
    if (f.display) {
      lines.push('\\[' + f.latex + '\\]');
    } else {
      lines.push('\\(' + f.latex + '\\)');
    }
  }
  lines.push('\\end{document}');
  return lines.join('\n');
}

function buildBibMap(blocks) {
  const map = {};
  if (!Array.isArray(blocks)) return map;
  blocks.forEach(block => {
    if (block && block.t === 'RawBlock' && block.c[0] === 'latex' && block.c[1].includes('\\begin{thebibliography}')) {
      const bibItems = block.c[1].match(/\\bibitem\{([^}]*)\}/g) || [];
      bibItems.forEach((item, i) => {
        const key = item.match(/\\bibitem\{([^}]*)\}/)[1];
        map[key] = i + 1;
      });
    }
  });
  return map;
}

function collectHeadings(blocks) {
  const headings = [];
  function walk(list) {
    if (!Array.isArray(list)) return;
    list.forEach(b => {
      if (!b) return;
      if (b.t === 'Header') {
        const level = b.c[0];
        const text = extractText(b.c[2]);
        if (text) headings.push({ level, text });
      }
      const sub = b.c ? (b.c[1] || b.c) : null;
      if (Array.isArray(sub)) walk(sub);
    });
  }
  walk(blocks);
  return headings;
}

function buildNumberedFormula(latex, sectNum, eqNum) {
  if (formulaTracker) {
    const idx = formulaTracker.counter++;
    formulaTracker.formulas.push({ index: idx, latex, display: true });
    return new DocxMath({
      children: [
        new MathRun(`__MATH_${idx}__`),
        new MathRun(`  #(${sectNum}-${eqNum})`)
      ]
    });
  }
  return new DocxMath({
    children: [
      new MathRun(latex),
      new MathRun(`  #(${sectNum}-${eqNum})`)
    ]
  });
}

function embedImage(src, options = {}) {
  if (!imageDir) return null;
  const imgPath = path.resolve(imageDir, src);
  if (!fs.existsSync(imgPath)) {
    return null;
  }
  const imgBuffer = fs.readFileSync(imgPath);
  const ext = path.extname(src).toLowerCase();
  return new ImageRun({
    data: imgBuffer,
    transformation: {
      width: options.width || 400,
      height: options.height || 300,
    },
    type: ext === '.png' ? ImageRun.PNG : ImageRun.JPEG,
    altText: options.alt || '',
  });
}

function buildDocument(ast, options = {}) {
  const {
    pageSize = 'A4',
    pageMargins = { top: 1440, bottom: 1440, left: 1440, right: 1440 },
    title = '',
    author = '',
    date = '',
    hasToc = false,
    formulaTracker: ft = null
  } = options;

  imageDir = options.inputDir || null;

  const pageConfig = {};
  if (pageSize === 'A4') {
    pageConfig.width = 11906;
    pageConfig.height = 16838;
  } else {
    pageConfig.width = 12240;
    pageConfig.height = 15840;
  }

  const docChildren = [];

  if (title) {
    docChildren.push(new Paragraph({ spacing: { before: 5000 }, children: [] }));
    docChildren.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: title, bold: true, size: 48, font: 'Times New Roman' })]
    }));
    if (author) {
      docChildren.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 },
        children: [new TextRun({ text: author, size: 28, font: 'Times New Roman' })]
      }));
    }
    if (date) {
      docChildren.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [new TextRun({ text: date, size: 24, font: 'Times New Roman' })]
      }));
    }
    docChildren.push(new Paragraph({ children: [new PageBreak()] }));
  }

  if (hasToc) {
    docChildren.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 400 },
      children: [new TextRun({ text: 'Contents', bold: true, size: 32, font: 'Times New Roman' })]
    }));
    docChildren.push(new TableOfContents('Table of Contents', {
      hyperlink: true,
      headingStyleRange: '1-3',
      useAppliedParagraphOutlineLevel: true
    }));
    docChildren.push(new Paragraph({ children: [new PageBreak()] }));
  }

  const blocks = ast.blocks || [];
  citationMap = buildBibMap(blocks);

  const prevTracker = formulaTracker;
  const prevFnCounter = footnoteCounter;
  const prevFnTexts = footnoteTexts;
  formulaTracker = ft;
  footnoteCounter = 0;
  footnoteTexts = {};
  eqCounter = 0;
  sectionNumber = 1;
  const mappedChildren = mapBlocks(blocks);
  formulaTracker = prevTracker;
  footnoteCounter = prevFnCounter;
  const footnotes = footnoteTexts;
  footnoteTexts = prevFnTexts;
  docChildren.push(...mappedChildren);

  return new Document({
    title: title || 'Converted Document',
    footnotes,
    styles: {
      default: {
        document: {
          run: { font: 'Times New Roman', size: 24 },
          paragraph: { spacing: { after: 120 } }
        }
      },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', run: { size: 32, bold: true }, paragraph: { spacing: { before: 240, after: 120 } }, outlineLevel: 0 },
        { id: 'Heading2', name: 'Heading 2', run: { size: 28, bold: true }, paragraph: { spacing: { before: 200, after: 100 } }, outlineLevel: 1 },
        { id: 'Heading3', name: 'Heading 3', run: { size: 26, bold: true }, paragraph: { spacing: { before: 160, after: 80 } }, outlineLevel: 2 },
        { id: 'Heading4', name: 'Heading 4', run: { size: 24, bold: true, italics: true }, paragraph: { spacing: { before: 120, after: 60 } }, outlineLevel: 3 },
      ]
    },
    numbering: {
      config: [
        {
          reference: 'bullets',
          levels: [
            { level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT }
          ]
        },
        {
          reference: 'numbers',
          levels: [
            { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT }
          ]
        }
      ]
    },
    sections: [{
      properties: {
        page: {
          size: pageConfig,
          margin: pageMargins
        }
      },
      ...(options.headerFooter ? {
        headers: options.headerFooter.headers,
        footers: options.headerFooter.footers
      } : {}),
      children: docChildren
    }]
  });
}

module.exports = {
  buildDocument, mapBlocks, mapBlock, mapInlines, mapInline, extractText,
  collectHeadings, formulasToTex, setFormulaTracker, buildBibMap
};
