const { TextRun, UnderlineType } = require('docx');

function parseBibText(latex) {
  latex = latex.replace(/\\end\{thebibliography\}.*$/, '').trim();
  latex = latex.replace(/[\n\r]+/g, ' ');
  const runs = []; let buf = '';
  function flush() { if (buf) { runs.push(new TextRun({ text: buf })); buf = ''; } }
  function parseArgs(i) {
    const args = []; let depth = 0;
    while (i < latex.length && latex[i] !== '{') i++;
    if (latex[i] === '{') { depth = 1; i++; }
    let arg = '';
    while (i < latex.length && depth > 0) {
      if (latex[i] === '{') { depth++; arg += '{'; i++; }
      else if (latex[i] === '}') { depth--; if (depth > 0) arg += '}'; i++; }
      else { arg += latex[i]; i++; }
    }
    if (arg) args.push(arg);
    while (latex[i] === '[') {
      i++; let opt = '';
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
          const cmd = m[1]; i += m[0].length;
          const { args, pos } = parseArgs(i); i = pos - 1;
          flush(); const content = args[0] || '';
          switch (cmd) {
            case 'textit': case 'emph': runs.push(new TextRun({ text: content, italics: true })); break;
            case 'textbf': runs.push(new TextRun({ text: content, bold: true })); break;
            case 'texttt': runs.push(new TextRun({ text: content, font: 'Courier New' })); break;
            case 'textsc': runs.push(new TextRun({ text: content, smallCaps: true })); break;
            case 'textcolor': {
              const color = args[0] || ''; const text = args[1] || args[0] || '';
              const codes = { red: 'FF0000', blue: '0000FF', green: '00AA00', gray: '888888' };
              runs.push(new TextRun({ text, color: codes[color] || color })); break;
            }
            case 'url': runs.push(new TextRun({ text: content, color: '0563C1', underline: { type: UnderlineType.SINGLE } })); break;
            default: if (content) runs.push(new TextRun({ text: content })); break;
          }
          continue;
        }
      }
      if (next === '{' || next === '}') { i += 2; continue; }
      if (next === '~' || next === ',' || next === ' ') { buf += ' '; i += 2; continue; }
      if (next === '\\') { buf += '\\'; i += 2; continue; }
      if (next === '`') { buf += '\u2018'; i += 2; continue; }
      if (next === '\'') { buf += '\u2019'; i += 2; continue; }
      if (['`', '\'', '^', '"', '~', '=', '.'].includes(next)) {
        const m = latex.slice(i).match(/^\\([`'\^"~=.])\{([^}]*)\}/);
        if (m) { buf += m[2]; i += m[0].length - 1; continue; }
      }
      if (/[&%$#_]/.test(next)) { buf += next; i += 2; continue; }
      i++; continue;
    }
    if (ch === '{' || ch === '}') continue;
    if (ch === '`' && latex[i + 1] === '`') { flush(); buf += '\u201C'; i++; continue; }
    if (ch === '\'' && latex[i + 1] === '\'') { flush(); buf += '\u201D'; i++; continue; }
    if (ch === '~') { buf += ' '; continue; }
    buf += ch;
  }
  flush();
  return runs;
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

function formulasToTex(formulas) {
  const lines = ['\\documentclass{article}', '\\usepackage{amsmath,amssymb}', '\\begin{document}'];
  for (const f of formulas) {
    if (f.display) lines.push('\\[' + f.latex + '\\]');
    else lines.push('\\(' + f.latex + '\\)');
  }
  lines.push('\\end{document}');
  return lines.join('\n');
}

module.exports = { parseBibText, buildBibMap, formulasToTex };
