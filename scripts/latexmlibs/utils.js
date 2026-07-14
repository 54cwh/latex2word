const { TextRun, ImageRun, UnderlineType } = require('docx');
const fs = require('fs');
const path = require('path');
const state = require('./state');
// NOT importing inline.js here to avoid circular dependency.
// expandFormatted lazily requires mapInline when called.

function extractText(inlines) {
  if (!inlines) return '';
  if (typeof inlines === 'string') return inlines;
  if (Array.isArray(inlines)) return inlines.map(i => extractText(i)).join('');
  if (typeof inlines === 'object') {
    if (inlines.t === 'Str') return inlines.c || '';
    if (inlines.t === 'Space' || inlines.t === 'SoftBreak') return ' ';
    if (inlines.t === 'Code' || inlines.t === 'Math') return inlines.c[1] || inlines.c || '';
    if (inlines.t === 'LineBreak') return ' ';
    if (inlines.t === 'RawInline') {
      if (inlines.c[0] === 'latex') return inlines.c[1].replace(/\\(label|ref|cite|pageref|index|glossary)\{[^}]*\}/g, '').replace(/[\u00A0~]/g, ' ').trim();
      return '';
    }
    if (inlines.c) return extractText(inlines.c);
  }
  return '';
}

function expandFormatted(items, fmt) {
  return items.flatMap(item => {
    const t = item.t; const c = item.c;
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
      default: return require('./inline').mapInline(item);
    }
  });
}

function embedImage(src, options = {}) {
  if (!state.imageDir) return null;
  const imgPath = path.resolve(state.imageDir, src);
  if (!fs.existsSync(imgPath)) return null;
  const imgBuffer = fs.readFileSync(imgPath);
  const ext = path.extname(src).toLowerCase();
  return new ImageRun({
    data: imgBuffer,
    transformation: { width: options.width || 400, height: options.height || 300 },
    type: ext === '.png' ? ImageRun.PNG : ImageRun.JPEG,
    altText: options.alt || '',
  });
}

function collectHeadings(blocks) {
  const headings = [];
  function walk(list) {
    if (!Array.isArray(list)) return;
    list.forEach(b => {
      if (!b) return;
      if (b.t === 'Header') { const text = extractText(b.c[2]); if (text) headings.push({ level: b.c[0], text }); }
      const sub = b.c ? (b.c[1] || b.c) : null;
      if (Array.isArray(sub)) walk(sub);
    });
  }
  walk(blocks);
  return headings;
}

module.exports = { extractText, expandFormatted, embedImage, collectHeadings };
