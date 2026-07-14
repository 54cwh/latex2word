const { TextRun, Paragraph, ExternalHyperlink, FootnoteReferenceRun, UnderlineType } = require('docx');
const { Math: DocxMath, MathRun } = require('docx');
const state = require('./state');
const { extractText, expandFormatted, embedImage } = require('./utils');

function mapInline(inline) {
  if (!inline) return [];
  if (Array.isArray(inline)) return inline.flatMap(i => mapInline(i));
  const t = inline.t; const c = inline.c;
  switch (t) {
    case 'Str':
      return [new TextRun(c.replace(/\u00A0/g, ' '))];
    case 'Space': return [new TextRun(' ')];
    case 'SoftBreak': return [new TextRun(' ')];
    case 'LineBreak': return [];
    case 'Strong': return expandFormatted(c, { bold: true });
    case 'Emph': return expandFormatted(c, { italics: true });
    case 'Underline': return expandFormatted(c, { underline: { type: UnderlineType.SINGLE } });
    case 'SmallCaps': return expandFormatted(c, { smallCaps: true });
    case 'Strikeout': return expandFormatted(c, { strike: true });
    case 'Superscript': return expandFormatted(c, { superScript: true });
    case 'Subscript': return expandFormatted(c, { subScript: true });
    case 'Code': return [new TextRun({ text: c[1], font: 'Courier New' })];
    case 'Math': {
      const mathType = c[0].t || c[0];
      const isDisplay = mathType === 'DisplayMath';
      const latex = c[1];
      if (state.formulaTracker) {
        const idx = state.formulaTracker.counter++;
        state.formulaTracker.formulas.push({ index: idx, latex, display: isDisplay });
        return [new DocxMath({ children: [new MathRun(`__MATH_${idx}__`)] })];
      }
      return [new DocxMath({ children: [new MathRun(latex)] })];
    }
    case 'Link': {
      const url = c[2][0];
      return [new ExternalHyperlink({
        children: [new TextRun({ text: extractText(c[1]), style: 'Hyperlink' })], link: url
      })];
    }
    case 'Image': {
      const src = c[2][0]; const alt = extractText(c[1]);
      const img = embedImage(src, { alt });
      if (img) return [img];
      return [new TextRun(`[Image: ${alt} — ${src}]`)];
    }
    case 'Note': {
      const noteId = ++state.footnoteCounter;
      const noteText = extractText(c[0] && c[0].c);
      state.footnoteTexts[noteId] = { children: [new Paragraph({ children: [new TextRun(noteText)] })] };
      return [new FootnoteReferenceRun(noteId)];
    }
    case 'Cite': {
      const citationId = (c[0] && c[0][0] && c[0][0].citationId) || '';
      const num = state.citationMap[citationId];
      if (num) return [new TextRun({ text: `[${num}]`, superScript: true })];
      return [new TextRun({ text: `[${citationId}]`, superScript: true })];
    }
    case 'Span': {
      const attrs = c[0] || [];
      const kvs = attrs[2] || [];
      const stylePair = kvs.find(kv => kv[0] === 'style');
      if (!stylePair) return mapInlines(c[1]);
      const styleVal = stylePair[1] || '';
      const colorMatch = styleVal.match(/color:\s*(\S+)/);
      if (colorMatch) {
        const color = colorMatch[1];
        const codes = { red: 'FF0000', blue: '0000FF', green: '00AA00', cyan: '00FFFF', magenta: 'FF00FF', yellow: 'FFFF00', black: '000000', white: 'FFFFFF', gray: '888888', orange: 'FF8000', purple: '800080', brown: '8B4513', pink: 'FFC0CB', olive: '808000', navy: '000080', teal: '008080' };
        const cc = codes[color.toLowerCase()] || color;
        return expandFormatted(c[1], { color: cc });
      }
      return mapInlines(c[1]);
    }
    case 'Quoted': return mapInlines(c[1]);
    case 'RawInline': {
      if (c[0] === 'latex') {
        let raw = c[1];
        raw = raw.replace(/\\(label|ref|cite|pageref|index|glossary)\{[^}]*\}/g, '');
        raw = raw.replace(/[\u00A0~]/g, ' ');
        const colorMatch = raw.match(/^\\textcolor\{([^}]*)\}\{((?:[^{}]|\{[^{}]*\})*)\}\s*/);
        if (colorMatch) {
          const color = colorMatch[1];
          const text = colorMatch[2];
          const codes = { red: 'FF0000', blue: '0000FF', green: '00AA00', cyan: '00FFFF', magenta: 'FF00FF', yellow: 'FFFF00', black: '000000', white: 'FFFFFF', gray: '888888', orange: 'FF8000', purple: '800080' };
          return [new TextRun({ text, color: codes[color] || color })];
        }
        const latexCmds = { texttrademark: '\u2122', textregistered: '\u00AE', textcopyright: '\u00A9', textbullet: '\u2022', textperiodcentered: '\u00B7', textemdash: '\u2014', textendash: '\u2013', textellipsis: '\u2026', textasciitilde: '~', textbackslash: '\\', textgreater: '>', textless: '<', textbar: '|', textdagger: '\u2020', textdaggerdbl: '\u2021', textsection: '\u00A7', textparagraph: '\u00B6', textsterling: '\u00A3', textyen: '\u00A5', texteuro: '\u20AC', textcent: '\u00A2', textdegree: '\u00B0', textpm: '\u00B1', texttimes: '\u00D7', textdiv: '\u00F7', micro: '\u00B5' };
        for (const [cmd, unicode] of Object.entries(latexCmds)) {
          raw = raw.replace(new RegExp(`\\\\${cmd}`, 'g'), unicode);
        }
        raw = raw.replace(/\\marginpar\{[^}]*\}/g, '');
        const cleaned = raw.trim();
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

function mapInlines(inlines) {
  if (!inlines) return [];
  return inlines.flatMap(inline => mapInline(inline));
}

module.exports = { mapInline, mapInlines };
