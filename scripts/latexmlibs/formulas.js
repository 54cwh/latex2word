const { Paragraph, Tab, TabStopType, AlignmentType } = require('docx');
const { Math: DocxMath, MathRun } = require('docx');
const state = require('./state');
const { mapInline, mapInlines } = require('./inline');

function buildNumberedFormula(latex, sectNum, eqNum) {
  if (state.formulaTracker) {
    const idx = state.formulaTracker.counter++;
    state.formulaTracker.formulas.push({ index: idx, latex, display: true });
    return new DocxMath({
      children: [new MathRun(`__MATH_${idx}__`), new MathRun(`  #(${sectNum}-${eqNum})`)]
    });
  }
  return new DocxMath({
    children: [new MathRun(latex), new MathRun(`  #(${sectNum}-${eqNum})`)]
  });
}

function buildParagraphs(inlines) {
  if (!inlines || !inlines.length) return [new Paragraph({ children: [] })];
  const segments = [];
  let current = [];
  for (const item of inlines) {
    if (item.t === 'Math' && (item.c[0].t || item.c[0]) === 'DisplayMath') {
      if (current.length) { segments.push({ items: current }); current = []; }
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
        const eqNum = ++state.eqCounter;
        return new Paragraph({
          tabStops: [
            { type: TabStopType.CENTER, position: 4513 },
            { type: TabStopType.RIGHT, position: 9026 },
          ],
          spacing: { before: 120, after: 120 },
          children: [
            new Tab(),
            buildNumberedFormula(latex, state.sectionNumber, eqNum),
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

module.exports = { buildParagraphs };
