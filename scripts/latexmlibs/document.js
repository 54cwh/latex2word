const { Document, Paragraph, TextRun, Header, Footer, PageNumber, AlignmentType, PageBreak, TableOfContents, LevelFormat } = require('docx');
const state = require('./state');
const { mapBlocks } = require('./blocks');
const { buildBibMap } = require('./bib');

function buildDocument(ast, options = {}) {
  const { pageSize = 'A4', pageMargins = { top: 1440, bottom: 1440, left: 1440, right: 1440 }, title = '', author = '', date = '', hasToc = false, formulaTracker: ft = null } = options;

  state.imageDir = options.inputDir || null;

  const pageConfig = {};
  if (pageSize === 'A4') { pageConfig.width = 11906; pageConfig.height = 16838; }
  else { pageConfig.width = 12240; pageConfig.height = 15840; }

  const docChildren = [];
  if (title) {
    docChildren.push(new Paragraph({ spacing: { before: 5000 }, children: [] }));
    docChildren.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: title, bold: true, size: 48, font: 'Times New Roman' })] }));
    if (author) docChildren.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: author, size: 28, font: 'Times New Roman' })] }));
    if (date) docChildren.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: date, size: 24, font: 'Times New Roman' })] }));
    docChildren.push(new Paragraph({ children: [new PageBreak()] }));
  }
  if (hasToc) {
    docChildren.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400, after: 400 }, children: [new TextRun({ text: 'Contents', bold: true, size: 32, font: 'Times New Roman' })] }));
    docChildren.push(new TableOfContents('Table of Contents', { hyperlink: true, headingStyleRange: '1-3', useAppliedParagraphOutlineLevel: true }));
    docChildren.push(new Paragraph({ children: [new PageBreak()] }));
  }

  const blocks = ast.blocks || [];
  state.citationMap = buildBibMap(blocks);
  const saved = state.saveState();
  state.formulaTracker = ft;
  state.footnoteCounter = 0;
  state.footnoteTexts = {};
  state.eqCounter = 0;
  state.sectionNumber = 1;
  const mappedChildren = mapBlocks(blocks);
  const footnotes = state.footnoteTexts;
  state.restoreState(saved);

  docChildren.push(...mappedChildren);

  return new Document({
    title: title || 'Converted Document',
    footnotes,
    styles: {
      default: {
        document: { run: { font: 'Times New Roman', size: 24 }, paragraph: { spacing: { after: 120 } } }
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
        { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT }] },
        { reference: 'numbers', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT }] }
      ]
    },
    sections: [{
      properties: { page: { size: pageConfig, margin: pageMargins } },
      ...(options.headerFooter ? { headers: options.headerFooter.headers, footers: options.headerFooter.footers } : {}),
      children: docChildren
    }]
  });
}

module.exports = { buildDocument };
