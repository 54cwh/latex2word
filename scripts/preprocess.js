const fs = require('fs');
const path = require('path');

function resolvePath(baseDir, texPath) {
  if (!texPath.endsWith('.tex')) texPath += '.tex';
  const resolved = path.resolve(baseDir, texPath);
  if (fs.existsSync(resolved)) return resolved;
  const withExt = texPath.endsWith('.tex') ? texPath : texPath + '.tex';
  const r2 = path.resolve(baseDir, withExt);
  if (fs.existsSync(r2)) return r2;
  return null;
}

function readTexFile(filePath) {
  return fs.readFileSync(filePath, 'utf-8');
}

function preprocess(mainTexPath) {
  const mainDir = path.dirname(path.resolve(mainTexPath));
  let tex = readTexFile(mainTexPath);

  const metadata = {
    hasMaketitle: false,
    hasToc: false,
    hasPrintbibliography: false,
    hasThebibliography: false,
    title: null,
    author: null,
    date: null,
    includes: [],
    headerFooter: null
  };

  const hf = extractHeaderFooter(tex);
  if (hf) metadata.headerFooter = hf;

  tex = tex.replace(/\\(documentclass)\s*(\[.*?\])?\s*(\{.*?\})/g, '');

  tex = tex.replace(/\\(title)\s*\{(.*?)\}/g, (m, cmd, content) => {
    metadata.title = content;
    return m;
  });

  tex = tex.replace(/\\(author)\s*\{(.*?)\}/g, (m, cmd, content) => {
    metadata.author = content.replace(/\sand\s/g, ', ');
    return m;
  });

  tex = tex.replace(/\\(date)\s*\{(.*?)\}/g, (m, cmd, content) => {
    metadata.date = content;
    return m;
  });

  let includeonly = [];
  tex = tex.replace(/\\(includeonly)\s*\{(.*?)\}/g, (m, cmd, content) => {
    includeonly = content.split(',').map(s => s.trim());
    return '';
  });

  tex = tex.replace(/\\(maketitle)/g, () => {
    metadata.hasMaketitle = true;
    return '\\par\\bigskip\\par';
  });

  tex = tex.replace(/\\(tableofcontents)/g, () => {
    metadata.hasToc = true;
    return '';
  });

  tex = tex.replace(/\\(printbibliography)\s*(\[.*?\])?/g, () => {
    metadata.hasPrintbibliography = true;
    return '';
  });

  const bibMatch = tex.match(/\\begin\{thebibliography\}/);
  if (bibMatch) metadata.hasThebibliography = true;

  const includePattern = /\\(?:input|include)(?![a-zA-Z])\s*(?:\{([^}]*)\}|(\S+))/g;
  let match;
  while ((match = includePattern.exec(tex)) !== null) {
    const incPath = match[1] || match[2];
    if (!incPath) continue;

    const incName = path.basename(incPath, '.tex');
    if (includeonly.length > 0 && !includeonly.some(io => io === incName || io === incPath)) {
      tex = tex.replace(match[0], '');
      continue;
    }

    const resolved = resolvePath(mainDir, incPath);
    if (resolved) {
      const content = readTexFile(resolved);
      const relPath = path.relative(mainDir, resolved);
      metadata.includes.push(relPath);
      tex = tex.replace(match[0], '\n' + content + '\n');
    } else {
      tex = tex.replace(match[0], `\n% Missing include: ${incPath}\n`);
    }
    includePattern.lastIndex = 0;
  }

  tex = expandNewcommands(tex);

  // Expand \makecell[align]{content} and \thead{content} for pandoc compatibility
  tex = expandCellCommands(tex);

  // Convert abstract environment to section heading + content for pandoc
  tex = tex.replace(/\\begin\{abstract\}([\s\S]*?)\\end\{abstract\}/g, (m, content) => {
    return '\\section*{Abstract}\n' + content.trim();
  });

  return { tex, metadata };
}

function extractHeaderFooter(tex) {
  const hf = { header: { left: '', center: '', right: '' }, footer: { left: '', center: '', right: '' } };

  const hasFancy = /\\usepackage\s*(?:\[.*?\])?\s*\{fancyhdr\}/.test(tex);
  const hasPagestyle = /\\pagestyle\s*\{fancy\}/.test(tex);
  if (!hasFancy && !hasPagestyle) return null;

  const simple = { lhead: ['header','left'], chead: ['header','center'], rhead: ['header','right'],
                   lfoot: ['footer','left'], cfoot: ['footer','center'], rfoot: ['footer','right'] };
  for (const [cmd, [sec, pos]] of Object.entries(simple)) {
    const re = new RegExp(`\\\\${cmd}\\s*\\{((?:[^{}]|\\{[^{}]*\\})*)\\}`, 'g');
    let m;
    while ((m = re.exec(tex)) !== null) hf[sec][pos] = expandPageRef(m[1]);
  }

  const fancyRe = /\\(fancyhead|fancyfoot)\s*\[(.*?)\]\s*\{((?:[^{}]|\{[^{}]*\})*)\}/g;
  let m;
  while ((m = fancyRe.exec(tex)) !== null) {
    const sec = m[1] === 'fancyhead' ? 'header' : 'footer';
    const pos = m[2];
    const content = expandPageRef(m[3]);
    if (pos.includes('L')) hf[sec].left = content;
    if (pos.includes('C')) hf[sec].center = content;
    if (pos.includes('R')) hf[sec].right = content;
  }

  return hf;
}

function expandPageRef(str) {
  return str
    .replace(/\\leftmark|\\rightmark|\\chaptermark/g, '')
    .replace(/\\thepage/g, '\\__PAGE_NUMBER__');
}

function expandNewcommands(tex) {
  const commands = {};

  const cmdDefPattern = /\\(re)?newcommand\s*(\*)?\s*\{(\w+)\}\s*\{((?:[^{}]|\{[^{}]*\})*)\}/g;
  tex = tex.replace(cmdDefPattern, (match, _re, _star, name, def) => {
    commands[name] = def;
    return '';
  });

  const opDefPattern = /\\DeclareMathOperator\s*\{(\w+)\}\s*\{([^}]*)\}/g;
  tex = tex.replace(opDefPattern, (match, name, opname) => {
    commands[name] = '\\operatorname{' + opname + '}';
    return '';
  });

  const sorted = Object.keys(commands).sort((a, b) => b.length - a.length);
  for (const cmd of sorted) {
    const re = new RegExp(`\\\\${cmd}(?![a-zA-Z])`, 'g');
    tex = tex.replace(re, commands[cmd]);
  }

  return tex;
}

function expandCellCommands(tex) {
  const regex = /\\(makecell|thead)(?![a-zA-Z])(?:\s*\[(.*?)\])?\s*\{/g;
  let result = '';
  let lastIndex = 0;
  let match;

  while ((match = regex.exec(tex)) !== null) {
    result += tex.slice(lastIndex, match.index);
    const start = match.index + match[0].length;
    const braceContent = extractBraces(tex, start);
    const inner = braceContent.content;
    const processed = inner.replace(/\\\\/g, '\\newline ');
    result += processed;
    lastIndex = braceContent.endIndex + 1;
  }

  result += tex.slice(lastIndex);
  return result;
}

function extractBraces(str, start) {
  let depth = 1;
  let i = start;
  while (i < str.length && depth > 0) {
    if (str[i] === '{') depth++;
    else if (str[i] === '}') depth--;
    if (depth > 0) i++;
  }
  return { content: str.slice(start, i), endIndex: i };
}

module.exports = { preprocess };
