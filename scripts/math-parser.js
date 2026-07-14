const { Math, MathRun } = require('docx');

function createMathElement(tex, isDisplay = false) {
  if (!tex || tex.trim() === '') return null;
  return new Math({ children: [new MathRun(tex)] });
}

module.exports = { createMathElement };
