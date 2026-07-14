const { Paragraph, TextRun, Header, Footer, PageNumber, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, PageBreak, LevelFormat } = require('docx');
const state = require('./state');
const { extractText, embedImage, expandFormatted } = require('./utils');
const { mapInline, mapInlines } = require('./inline');
const { buildParagraphs } = require('./formulas');
const { parseBibText } = require('./bib');

function mapBlock(block) {
  if (!block) return [];
  const t = block.t; const c = block.c;
  switch (t) {
    case 'Header': {
      const level = c[0];
      if (level === 1) { state.sectionNumber++; state.eqCounter = 0; }
      return [new Paragraph({ heading: HeadingLevel[`HEADING_${level}`], children: mapInlines(c[2]) })];
    }
    case 'Para': return buildParagraphs(c);
    case 'Plain': return buildParagraphs(c);
    case 'BulletList': {
      const items = [];
      c.forEach(itemBlocks => {
        const firstBlock = itemBlocks[0];
        if (firstBlock && firstBlock.t === 'Para') {
          items.push(new Paragraph({ numbering: { reference: 'bullets', level: 0 }, children: mapInlines(firstBlock.c) }));
        }
      });
      return items;
    }
    case 'OrderedList': {
      const items = [];
      c[1].forEach(itemBlocksInner => {
        const firstBlock = itemBlocksInner[0];
        if (firstBlock && firstBlock.t === 'Para') {
          items.push(new Paragraph({ numbering: { reference: 'numbers', level: 0 }, children: mapInlines(firstBlock.c) }));
        }
      });
      return items;
    }
    case 'Table': {
      const colspecs = c[2] || []; const thead = c[3] || []; const tbodies = c[4] || [];
      const cols = colspecs.length || 1; const colWidth = Math.floor(9026 / cols);
      function cellFromPandoc(cell, isHeader) {
        const cellBlocks = cell[4] || [];
        const opts = {
          columnSpan: cell[3] || 1, rowSpan: cell[2] || 1,
          width: { size: colWidth * (cell[3] || 1), type: WidthType.DXA },
          children: cellBlocks.flatMap(b => mapBlock(b)),
        };
        if (!opts.children.length) opts.children = [new Paragraph({ children: [] })];
        if (isHeader) opts.borders = { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000' } };
        return new TableCell(opts);
      }
      function rowFromPandoc(rowData, isHeader) {
        return new TableRow({ children: (rowData[1] || []).map(cell => cellFromPandoc(cell, isHeader)) });
      }
      const rows = [];
      if (thead.length >= 2 && Array.isArray(thead[1])) thead[1].forEach(row => rows.push(rowFromPandoc(row, true)));
      tbodies.forEach(tbody => {
        const bodyRows = Array.isArray(tbody[3]) ? tbody[3] : [];
        bodyRows.forEach(row => { if (Array.isArray(row) && row.length >= 2 && Array.isArray(row[1])) rows.push(rowFromPandoc(row)); });
        if (Array.isArray(tbody[2])) tbody[2].forEach(row => { if (Array.isArray(row) && row.length >= 2 && Array.isArray(row[1])) rows.push(rowFromPandoc(row, false)); });
      });
      const result = [];
      if (c[1] && c[1][1]) {
        const ct = extractText(c[1][1]);
        if (ct) result.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120 }, children: [new TextRun({ text: ct, bold: true, italics: true, size: 20 })] }));
      }
      result.push(new Table({
        width: { size: 9026, type: WidthType.DXA }, columnWidths: Array(cols).fill(colWidth),
        borders: { top: { style: BorderStyle.SINGLE, size: 12, color: '000000' }, bottom: { style: BorderStyle.SINGLE, size: 12, color: '000000' }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
        rows
      }));
      result.push(new Paragraph({ children: [new TextRun('')] }));
      return result;
    }
    case 'CodeBlock': {
      const codeText = c[1]; const lang = (c[0] && c[0][1] && c[0][1][0]) || '';
      const result = [];
      if (lang) result.push(new Paragraph({ spacing: { before: 120 }, children: [new TextRun({ text: `[${lang}]`, size: 18, color: '666666', italics: true })] }));
      codeText.split('\n').forEach(line => result.push(new Paragraph({ spacing: { before: 0, after: 0 }, shading: { type: ShadingType.CLEAR, fill: 'F0F0F0' }, indent: { left: 360 }, children: [new TextRun({ text: line, font: 'Courier New', size: 18 })] })));
      result.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
      return result;
    }
    case 'BlockQuote': {
      return c.flatMap(b => mapBlock(b).map(p => p instanceof Paragraph ? new Paragraph({ indent: { left: 720, right: 720 }, spacing: { before: 80, after: 80 }, children: p.children }) : p));
    }
    case 'HorizontalRule': {
      return [new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '999999', space: 1 } }, spacing: { before: 200, after: 200 } })];
    }
    case 'Div': {
      const classes = c[0][1] || []; const divBlocks = c[1];
      if (classes.includes('theorem')) {
        const contentBlocks = divBlocks.flatMap(b => mapBlock(b));
        return [new Paragraph({ spacing: { before: 120, after: 60 }, children: [new TextRun({ text: 'Theorem.', bold: true, italics: true }), ...(contentBlocks[0] ? contentBlocks[0].children || [] : [])] }), ...contentBlocks.slice(1)];
      }
      if (classes.includes('proof')) {
        const contentBlocks = divBlocks.flatMap(b => mapBlock(b));
        if (!contentBlocks.length) return [];
        const firstChildren = contentBlocks[0] && contentBlocks[0].children ? contentBlocks[0].children : [];
        const proofFirst = new Paragraph({ spacing: { before: 120, after: 60 }, children: [new TextRun({ text: 'Proof.', italics: true }), new TextRun(' '), ...firstChildren] });
        const rest = contentBlocks.slice(1).map((b, i, arr) => i === arr.length - 1 && b.children ? new Paragraph({ spacing: { before: 60, after: 60 }, children: [...b.children, new TextRun(' □')] }) : b);
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
            alignment: AlignmentType.LEFT,
            indent: { left: 360, hanging: 360 },
            spacing: { before: 60, after: 60 },
            children: [new TextRun({ text: `[${i + 1}] ` }), ...parseBibText(text)]
          });
        });
        return [new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('Bibliography')] }), ...items];
      }
      return [];
    }
    case 'Figure': {
      const figBlocks = c[2]; const caption = c[1]; const capText = extractText(caption);
      const result = [];
      figBlocks.forEach(b => {
        if (b.t === 'Image') {
          const src = b.c[2][0]; const alt = extractText(b.c[1]);
          const img = embedImage(src, { alt, width: 500 });
          if (img) result.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [img] }));
          else result.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun(`[Figure: ${capText}] (${src})`)] }));
        } else {
          result.push(...mapBlock(b));
        }
      });
      if (capText) result.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60 }, children: [new TextRun({ text: capText, italics: true, size: 20 })] }));
      return result;
    }
    case 'LineBlock': return [new Paragraph({ children: mapInlines(c) })];
    default: return [];
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
    if (items.t === 'Math') formulas.push({ latex: items.c[1], display: (items.c[0].t || items.c[0]) === 'DisplayMath' });
    if (items.c) { if (Array.isArray(items.c)) items.c.forEach(walkInlines); else walkInlines(items.c); }
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
        case 'Table': [b.c[3], b.c[4]].forEach(part => { if (Array.isArray(part)) part.forEach(row => { if (Array.isArray(row)) (Array.isArray(row[1]) ? row[1] : row).forEach(cell => { if (cell && cell[4]) walkBlocks(cell[4]); }); }); }); break;
      }
    });
  }
  walkBlocks(Array.isArray(blocks) ? blocks : (blocks.blocks || []));
  return formulas;
}

module.exports = { mapBlock, mapBlocks, collectFormulasFromAst };
