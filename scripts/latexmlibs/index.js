const { buildDocument } = require('./document');
const { mapBlocks, mapBlock, collectFormulasFromAst } = require('./blocks');
const { mapInline, mapInlines } = require('./inline');
const { extractText, collectHeadings } = require('./utils');
const { formulasToTex } = require('./bib');
const { setFormulaTracker } = require('./state');

module.exports = {
  buildDocument, mapBlocks, mapBlock, mapInlines, mapInline, extractText,
  collectHeadings, formulasToTex, setFormulaTracker, collectFormulasFromAst, buildBibMap: require('./bib').buildBibMap
};
