const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const { preprocess } = require('./preprocess');
const { buildDocument, extractText, formulasToTex } = require('./latexmlibs');

function checkPandoc() {
  try {
    const v = execSync('pandoc --version', { encoding: 'utf-8' });
    const match = v.match(/pandoc\s+([\d.]+)/);
    if (match) console.error(`pandoc ${match[1]} detected`);
    return true;
  } catch (e) {
    console.error('Error: pandoc is required but not found.');
    console.error('Install: https://pandoc.org/installing.html');
    process.exit(1);
  }
}

function extractMetadataFromAst(ast) {
  const meta = ast.meta || {};
  const title = extractText(meta.title && meta.title.c);
  let author = '';
  if (meta.author) {
    if (meta.author.t === 'MetaList') {
      author = meta.author.c.map(a => extractText(a.c)).join('; ');
    } else {
      author = extractText(meta.author.c);
    }
  }
  const date = extractText(meta.date && meta.date.c);
  return { title, author, date };
}

function runPandoc(texContent) {
  const tmpDir = fs.mkdtempSync(path.join(require('os').tmpdir(), 'l2w-'));
  const tmpTex = path.join(tmpDir, 'input.tex');
  fs.writeFileSync(tmpTex, texContent, 'utf-8');

  const json = execSync(
    `pandoc -f latex+raw_tex -t json "${tmpTex}" 2>/dev/null`,
    { encoding: 'utf-8', maxBuffer: 50 * 1024 * 1024, timeout: 30000 }
  );

  const ast = JSON.parse(json);
  fs.rmSync(tmpDir, { recursive: true, force: true });
  return ast;
}

async function latex2docx(inputPath, outputPath) {
  checkPandoc();

  const absInput = path.resolve(inputPath);
  if (!fs.existsSync(absInput)) {
    console.error(`Error: input file not found: ${absInput}`);
    process.exit(1);
  }

  console.error(`Reading: ${absInput}`);

  const { tex, metadata } = preprocess(absInput);

  console.error(`Running pandoc...`);
  const ast = runPandoc(tex);

  const metaInfo = extractMetadataFromAst(ast);
  const title = metaInfo.title || metadata.title || '';
  const author = metaInfo.author || metadata.author || '';
  const date = metaInfo.date || metadata.date || '';
  const hasMaketitle = metadata.hasMaketitle || !!(title);

  console.error(`Title: ${title || '(none)'}`);
  if (metadata.includes.length > 0) {
    console.error(`Merged files: ${metadata.includes.join(', ')}`);
  }
  if (metadata.hasToc) console.error('Table of Contents detected');
  if (metadata.hasPrintbibliography || metadata.hasThebibliography) console.error('Bibliography detected');

  const formulaTracker = { formulas: [], counter: 0 };

  const headerFooter = metadata.headerFooter ? buildHeaderFooter(metadata.headerFooter) : null;

  const doc = buildDocument(ast, {
    title: hasMaketitle ? title : '',
    author,
    date,
    hasToc: metadata.hasToc,
    formulaTracker,
    headerFooter,
    inputDir: path.dirname(absInput)
  });

  const formulas = formulaTracker.formulas;
  console.error(`Formulas: ${formulas.length} collected`);

  const outputAbs = path.resolve(outputPath);
  const outputDir = path.dirname(outputAbs);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const tempDocx = outputAbs + '.tmp';
  const Packer = require('docx').Packer;
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(tempDocx, buffer);
  console.error(`Placeholder docx: ${(buffer.length / 1024).toFixed(1)} KB`);

  if (formulas.length > 0) {
    const formulasTex = formulasToTex(formulas);
    const tmpDir = fs.mkdtempSync('/tmp/l2w-omml-');
    const texPath = path.join(tmpDir, 'formulas.tex');
    fs.writeFileSync(texPath, formulasTex, 'utf-8');
    const scriptDir = __dirname;
    execSync(
      `python3 "${scriptDir}/inject-omml.py" "${tempDocx}" "${outputAbs}" --formulas-tex "${texPath}"`,
      { stdio: 'inherit', timeout: 60000 }
    );
    fs.rmSync(tmpDir, { recursive: true, force: true });
    fs.rmSync(tempDocx);
    const finalSize = fs.statSync(outputAbs).size;
    console.error(`\nOutput: ${outputAbs} (${(finalSize / 1024).toFixed(1)} KB)`);
  } else {
    fs.renameSync(tempDocx, outputAbs);
    console.error(`\nOutput: ${outputAbs} (${(buffer.length / 1024).toFixed(1)} KB, no formulas)`);
  }
}

function parseHFContent(str) {
  const { TextRun, PageNumber } = require('docx');
  const parts = str.split(/\\__PAGE_NUMBER__/);
  if (parts.length === 1) return buildHFRuns(str);
  const children = [];
  for (let i = 0; i < parts.length; i++) {
    if (parts[i]) children.push(...buildHFRuns(parts[i]));
    if (i < parts.length - 1) children.push(new TextRun({ children: [PageNumber.CURRENT], size: 18 }));
  }
  return children;
}

function buildHFRuns(str) {
  const { TextRun } = require('docx');
  const runs = [];
  let i = 0;
  let plain = '';
  function flushPlain() {
    if (plain) { runs.push(new TextRun({ text: plain, size: 18 })); plain = ''; }
  }
  while (i < str.length) {
    const rest = str.slice(i);
    const boldCmd = rest.match(/^\\textbf\s*\{((?:[^{}]|\{[^{}]*\})*)\}/);
    if (boldCmd) { flushPlain(); runs.push(new TextRun({ text: boldCmd[1], bold: true, size: 18 })); i += boldCmd[0].length; continue; }
    const italicCmd = rest.match(/^\\(?:textit|emph)\s*\{((?:[^{}]|\{[^{}]*\})*)\}/);
    if (italicCmd) { flushPlain(); runs.push(new TextRun({ text: italicCmd[1], italics: true, size: 18 })); i += italicCmd[0].length; continue; }
    const ttCmd = rest.match(/^\\texttt\s*\{((?:[^{}]|\{[^{}]*\})*)\}/);
    if (ttCmd) { flushPlain(); runs.push(new TextRun({ text: ttCmd[1], font: 'Courier New', size: 18 })); i += ttCmd[0].length; continue; }
    const scCmd = rest.match(/^\\textsc\s*\{((?:[^{}]|\{[^{}]*\})*)\}/);
    if (scCmd) { flushPlain(); runs.push(new TextRun({ text: scCmd[1].toUpperCase(), font: 'Times New Roman', size: 18 })); i += scCmd[0].length; continue; }
    if (rest.startsWith('---')) { plain += '\u2014'; i += 3; continue; }
    if (rest.startsWith('--')) { plain += '\u2013'; i += 2; continue; }
    plain += str[i];
    i++;
  }
  flushPlain();
  return runs;
}

function buildHeaderFooter(hf) {
  const { Header, Footer, Paragraph, AlignmentType } = require('docx');
  const result = {};
  for (const [key, Cls, posKey] of [['header', Header, 'headers'], ['footer', Footer, 'footers']]) {
    const data = hf[key];
    const paragraphs = [];
    if (data.left) paragraphs.push(new Paragraph({ alignment: AlignmentType.LEFT, children: parseHFContent(data.left) }));
    if (data.center) paragraphs.push(new Paragraph({ alignment: AlignmentType.CENTER, children: parseHFContent(data.center) }));
    if (data.right) paragraphs.push(new Paragraph({ alignment: AlignmentType.RIGHT, children: parseHFContent(data.right) }));
    if (paragraphs.length) result[posKey] = { default: new Cls({ children: paragraphs }) };
  }
  return result;
}

if (require.main === module) {
  const args = process.argv.slice(2);
  if (args.length < 1) {
    console.error('Usage: node latex2docx.js <input.tex> [output.docx]');
    process.exit(1);
  }
  const input = args[0];
  const output = args[1] || input.replace(/\.tex$/, '.docx');
  latex2docx(input, output).catch(err => {
    console.error('Fatal:', err);
    process.exit(1);
  });
}

module.exports = { latex2docx };
