'use strict';

const { stripTrailingComment } = require('./index');

const CHR_RE = /\b(Chr[WB]?\s*\$?)\s*\(\s*(&H[0-9A-Fa-f]+|\d+)\s*\)/gi;

const NAMED = {
  0: { glyph: '␀', name: 'NUL' },
  7: { glyph: '🔔', name: 'Bell' },
  8: { glyph: '⌫', name: 'Backspace' },
  9: { glyph: '⇥', name: 'Tab' },
  10: { glyph: '↵', name: 'Line feed' },
  11: { glyph: '⤓', name: 'Vertical tab' },
  12: { glyph: '⇞', name: 'Form feed' },
  13: { glyph: '⏎', name: 'Return' },
  27: { glyph: '⎋', name: 'Escape' },
  32: { glyph: '␣', name: 'Space' },
  127: { glyph: '⌦', name: 'Delete' },
  160: { glyph: '⍽', name: 'Non-breaking space' }
};

function parseCode(raw) {
  if (!raw) return null;
  const text = String(raw).trim();
  if (/^&H/i.test(text)) {
    const n = parseInt(text.slice(2), 16);
    return Number.isFinite(n) ? n : null;
  }
  if (/^\d+$/.test(text)) return Number(text);
  return null;
}

function isWide(fnName) {
  return /^ChrW/i.test(fnName);
}

function inRange(code, wide) {
  if (!Number.isInteger(code) || code < 0) return false;
  return wide ? code <= 65535 : code <= 255;
}

function isPrintable(code) {
  if (code >= 33 && code <= 126) return true;
  if (code === 32 || code === 160) return false;
  if (code < 32 || code === 127) return false;
  try {
    const ch = String.fromCodePoint(code);
    return ch.length > 0 && /[\p{L}\p{N}\p{P}\p{S}\p{Sm}\p{Sc}]/u.test(ch);
  } catch (_) {
    return false;
  }
}

function describe(code, fnName) {
  if (!inRange(code, isWide(fnName || 'Chr'))) return null;
  const named = NAMED[code];
  if (named) {
    return {
      code,
      glyph: named.glyph,
      name: named.name,
      label: named.glyph,
      tooltip: named.name + ' (Chr ' + code + ')',
      markdown: '**' + named.glyph + '**  ' + named.name
    };
  }
  let glyph = '';
  try {
    glyph = String.fromCodePoint(code);
  } catch (_) {
    return null;
  }
  if (!isPrintable(code)) {
    const hex = 'U+' + code.toString(16).toUpperCase().padStart(4, '0');
    return {
      code,
      glyph: hex,
      name: hex,
      label: hex,
      tooltip: hex,
      markdown: '`' + hex + '`'
    };
  }
  return {
    code,
    glyph,
    name: glyph,
    label: glyph,
    tooltip: '"' + glyph + '" (Chr ' + code + ')',
    markdown: '**' + glyph + '**'
  };
}

function findCalls(text) {
  const out = [];
  const lines = String(text || '').split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const codeLine = stripTrailingComment(lines[i]);
    CHR_RE.lastIndex = 0;
    let m;
    while ((m = CHR_RE.exec(codeLine))) {
      const fnName = m[1];
      const code = parseCode(m[2]);
      const info = describe(code, fnName);
      if (!info) continue;
      out.push({
        line: i,
        start: m.index,
        end: m.index + m[0].length,
        fnName,
        ...info
      });
    }
  }
  return out;
}

function previewAt(lineText, word) {
  if (!word || !/^Chr[WB]?\$?$/i.test(word.name)) return null;
  const fromWord = lineText.slice(word.start);
  CHR_RE.lastIndex = 0;
  const m = CHR_RE.exec(fromWord);
  if (!m || m.index !== 0) return null;
  return describe(parseCode(m[2]), m[1]);
}

function appendHover(text, lineText, word) {
  const info = previewAt(lineText, word);
  if (!info) return text;
  const extra = '\n\n' + info.markdown;
  if (!text) return info.markdown;
  if (text.indexOf(info.markdown) >= 0) return text;
  return text + extra;
}

module.exports = {
  parseCode,
  describe,
  findCalls,
  previewAt,
  appendHover
};
