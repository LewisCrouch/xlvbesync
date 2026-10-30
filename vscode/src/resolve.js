'use strict';

const { enclosingWith, enclosingProc, varsForLine } = require('./index');
const builtins = require('./builtins');
const chrPreview = require('./chrPreview');

const IDENT_RE = /[A-Za-z_][A-Za-z0-9_]*[\$%&!#@]?/g;

function wordAt(lineText, character) {
  IDENT_RE.lastIndex = 0;
  let m;
  while ((m = IDENT_RE.exec(lineText))) {
    if (character >= m.index && character <= m.index + m[0].length) {
      return { name: m[0], start: m.index, end: m.index + m[0].length };
    }
  }
  return null;
}

function qualifierBefore(lineText, wordStart) {
  const before = lineText.slice(0, wordStart);
  const m = before.match(/([A-Za-z_][A-Za-z0-9_]*)\s*\.\s*$/);
  return m ? m[1] : '';
}

function memberTrigger(lineText, character) {
  const before = lineText.slice(0, character);
  const qualified = before.match(/([A-Za-z_][A-Za-z0-9_]*)\s*\.\s*([A-Za-z_][A-Za-z0-9_]*)?$/);
  if (qualified) {
    return { qualifier: qualified[1], partial: qualified[2] || '', fromDot: false };
  }
  if (/\.\s*([A-Za-z_][A-Za-z0-9_]*)?$/.test(before) && /^\s*\./.test(before.trimStart ? before.replace(/^\s+/, '') : before.replace(/^\s+/, ''))) {
    const partialM = before.match(/\.([A-Za-z_][A-Za-z0-9_]*)?$/);
    return { qualifier: '', partial: (partialM && partialM[1]) || '', fromDot: true };
  }
  return null;
}

function uniqueMembers(mod, { includePrivate }) {
  const seen = new Map();
  for (const mem of mod.members) {
    if (mem.isPrivate && !includePrivate) continue;
    const key = mem.name.toLowerCase();
    if (!seen.has(key)) seen.set(key, mem);
  }
  return [...seen.values()];
}

function findMember(mod, name, { includePrivate } = { includePrivate: true }) {
  const lower = name.toLowerCase();
  return mod.members.filter((m) => m.name.toLowerCase() === lower && (includePrivate || !m.isPrivate));
}

function resolveTypeName(index, currentPath, qualifier, documentText, lineNumber) {
  if (!qualifier) return '';
  if (/^me$/i.test(qualifier)) {
    const here = index.moduleAt(currentPath);
    return here ? here.name : '';
  }

  const named = index.findModulesByName(qualifier, currentPath);
  if (named.length) return named[0].name;

  const globals = index.findGlobals(qualifier, currentPath);
  if (globals.length) return globals[0].global.typeName;

  const here = index.moduleAt(currentPath);
  if (here) {
    const vars = varsForLine(here, lineNumber);
    const typed = vars.get(qualifier.toLowerCase());
    if (typed) return typed;
    const members = findMember(here, qualifier, { includePrivate: true });
    if (members.length && members[0].returnType) return members[0].returnType;
  }

  const g = builtins.findGlobal(qualifier);
  if (g && g.typeName) return g.typeName;
  if (builtins.findType(qualifier)) return qualifier;

  return '';
}

function excelTypeForDocument(mod) {
  if (!mod || mod.kind !== 'document') return '';
  if (/^ThisWorkbook$/i.test(mod.name)) return 'Workbook';
  if (/^Sheet/i.test(mod.name)) return 'Worksheet';
  return '';
}

function builtinModule(typeName) {
  return builtins.asModule(builtins.findType(typeName));
}

function resolvedTypeBundle(index, currentPath, qualifier, documentText, lineNumber) {
  const typeName = resolveTypeName(index, currentPath, qualifier, documentText, lineNumber);
  if (!typeName) return null;
  const proj = moduleForType(index, typeName, currentPath);
  if (proj) {
    return {
      typeName,
      mod: proj,
      builtinMod: builtinModule(excelTypeForDocument(proj))
    };
  }
  const builtin = builtinModule(typeName);
  if (!builtin) return null;
  return { typeName, mod: builtin, builtinMod: builtin };
}

function moduleForType(index, typeName, currentPath) {
  if (!typeName) return null;
  const mods = index.findModulesByName(typeName, currentPath);
  return mods[0] || null;
}

function paramAsMember(proc, p) {
  const typeBit = p.type ? ' As ' + p.type : '';
  return {
    name: p.name,
    kind: 'Parameter',
    access: 'Private',
    isPrivate: true,
    signature: p.name + typeBit,
    line: proc.line,
    returnType: p.type || '',
    declParams: [],
    doc: { summary: p.text || '', params: [], returns: '' },
    params: []
  };
}

function localDecl(mod, lineNumber, name) {
  const proc = enclosingProc(mod, lineNumber);
  if (!proc) return null;
  const lower = name.toLowerCase();
  const param = (proc.params || []).find((p) => p.name.toLowerCase() === lower);
  if (param) return { hostName: proc.name, mem: paramAsMember(proc, param) };
  if (proc.locals && proc.locals.has(lower)) {
    return { hostName: proc.name, mem: proc.locals.get(lower) };
  }
  return null;
}

function definitionsAt(index, currentPath, lineText, lineNumber, character, documentText) {
  const word = wordAt(lineText, character);
  if (!word) return [];

  const qualifier = qualifierBefore(lineText, word.start);
  if (qualifier) {
    const typeName = resolveTypeName(index, currentPath, qualifier, documentText, lineNumber);
    const mod = moduleForType(index, typeName, currentPath);
    if (!mod) return [];
    return findMember(mod, word.name, { includePrivate: /^me$/i.test(qualifier) }).map((mem) => ({
      fsPath: mod.fsPath,
      line: mem.line,
      name: mem.name
    }));
  }

  const named = index.findModulesByName(word.name, currentPath);
  if (named.length) {
    return [{ fsPath: named[0].fsPath, line: 0, name: named[0].name }];
  }

  const here = index.moduleAt(currentPath);
  if (here) {
    const local = findMember(here, word.name, { includePrivate: true });
    if (local.length) {
      return local.map((mem) => ({ fsPath: here.fsPath, line: mem.line, name: mem.name }));
    }
    const decl = localDecl(here, lineNumber, word.name);
    if (decl) {
      return [{ fsPath: here.fsPath, line: decl.mem.line, name: decl.mem.name }];
    }
    const vars = varsForLine(here, lineNumber);
    const typed = vars.get(word.name.toLowerCase());
    if (typed) {
      const typeMod = moduleForType(index, typed, currentPath);
      if (typeMod) return [{ fsPath: typeMod.fsPath, line: 0, name: typeMod.name }];
    }
  }

  const globals = index.findGlobals(word.name, currentPath);
  if (globals.length) {
    const typeMod = moduleForType(index, globals[0].global.typeName, currentPath);
    if (typeMod) return [{ fsPath: typeMod.fsPath, line: 0, name: typeMod.name }];
  }

  return [];
}

function completionsAt(index, currentPath, lineText, lineNumber, character, documentText) {
  const trigger = memberTrigger(lineText, character);
  let qualifier = '';
  let includePrivate = false;

  if (trigger && trigger.qualifier) {
    qualifier = trigger.qualifier;
  } else if (trigger && trigger.fromDot) {
    qualifier = enclosingWith(documentText, lineNumber) || 'Me';
    includePrivate = /^me$/i.test(qualifier);
  } else {
    const word = wordAt(lineText, character);
    const q = word ? qualifierBefore(lineText, word.start) : '';
    if (q) qualifier = q;
  }

  if (qualifier) {
    includePrivate = includePrivate || /^me$/i.test(qualifier);
    const bundle = resolvedTypeBundle(index, currentPath, qualifier, documentText, lineNumber);
    if (!bundle) return [];
    const items = [];
    const seen = new Set();
    const pushFrom = (mod, priv) => {
      if (!mod) return;
      for (const mem of uniqueMembers(mod, { includePrivate: priv })) {
        const key = mem.name.toLowerCase();
        if (seen.has(key)) continue;
        seen.add(key);
        items.push({
          name: mem.name,
          kind: mem.kind,
          signature: mem.signature,
          detail: mod.name,
          documentation: formatMemberHover(mod, [mem])
        });
      }
    };
    if (bundle.mod && !bundle.mod.isBuiltin) pushFrom(bundle.mod, includePrivate);
    pushFrom(bundle.mod && bundle.mod.isBuiltin ? bundle.mod : bundle.builtinMod, true);
    return items;
  }

  const items = [];
  const seen = new Set();
  for (const mod of index.preferred(index.all(), currentPath)) {
    const key = mod.name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    items.push({
      name: mod.name,
      kind: mod.kind === 'class' || mod.kind === 'form' ? 'Class' : 'Module',
      signature: mod.kind + ' ' + mod.name,
      detail: mod.subtree || mod.kind,
      documentation: formatModuleHover(mod)
    });
  }
  for (const g of builtins.getCatalog().globals.values()) {
    const key = g.name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    const owner = g.owner || 'VBA';
    items.push({
      name: g.name,
      kind: g.member ? g.member.kind : 'Class',
      signature: g.member ? g.member.signature : g.name,
      detail: owner,
      documentation: g.member
        ? formatMemberHover({ name: owner, members: [g.member] }, [g.member])
        : formatModuleHover({ name: g.name, kind: 'class', subtree: owner })
    });
  }
  return items;
}

function memberWithDocs(mod, mem) {
  if (mem.doc && mem.doc.summary) return mem;
  if (!/^property (let|set)$/i.test(mem.kind)) return mem;
  const getter = mod.members.find((m) => (
    m.name.toLowerCase() === mem.name.toLowerCase()
    && /^property get$/i.test(m.kind)
    && m.doc
    && m.doc.summary
  ));
  if (!getter) return mem;
  return Object.assign({}, mem, {
    doc: getter.doc,
    params: (mem.params && mem.params.length) ? mem.params : getter.params
  });
}

function formatMemberHover(mod, mems) {
  if (!mems.length) return '';
  const mem = memberWithDocs(mod, mems[0]);
  const lines = [];
  lines.push('**' + mod.name + '.' + mem.name + '**');
  lines.push('');
  if (mem.doc && mem.doc.summary) {
    lines.push(mem.doc.summary);
    lines.push('');
  }
  lines.push('```vba');
  for (const m of mems) lines.push(m.signature);
  lines.push('```');
  if (mem.params && mem.params.length) {
    lines.push('');
    lines.push('**Parameters**');
    for (const p of mem.params) {
      const bits = [];
      if (p.paramArray) bits.push('ParamArray');
      if (p.type) bits.push('`' + p.type + '`');
      if (p.optional) bits.push(p.defaultValue ? 'optional, default `' + p.defaultValue + '`' : 'optional');
      const typeBit = bits.length ? ' (' + bits.join(', ') + ')' : '';
      const desc = p.text ? ' — ' + p.text : '';
      lines.push('- `' + p.name + '`' + typeBit + desc);
    }
  }
  const returnsText = (mem.doc && mem.doc.returns) || (mem.returnType ? '`' + mem.returnType + '`' : '');
  if (returnsText && !/^(sub|parameter|const|variable)$/i.test(mem.kind)) {
    lines.push('');
    lines.push('**Returns**');
    lines.push(returnsText);
  }
  return lines.join('\n').trim();
}

function formatModuleHover(mod) {
  return '**' + mod.name + '**  \n' + mod.kind + (mod.subtree ? ' (`' + mod.subtree + '`)' : '');
}

function hoverAt(index, currentPath, lineText, lineNumber, character, documentText) {
  const word = wordAt(lineText, character);
  if (!word) return null;

  let text = null;
  const qualifier = qualifierBefore(lineText, word.start);
  if (qualifier) {
    const bundle = resolvedTypeBundle(index, currentPath, qualifier, documentText, lineNumber);
    if (!bundle) return null;
    if (bundle.mod && !bundle.mod.isBuiltin) {
      const local = findMember(bundle.mod, word.name, { includePrivate: /^me$/i.test(qualifier) });
      if (local.length) text = formatMemberHover(bundle.mod, local);
    }
    if (!text) {
      const builtinHost = bundle.mod && bundle.mod.isBuiltin ? bundle.mod : bundle.builtinMod;
      if (builtinHost) {
        const mems = findMember(builtinHost, word.name, { includePrivate: true });
        if (mems.length) text = formatMemberHover(builtinHost, mems);
      }
    }
    return chrPreview.appendHover(text, lineText, word);
  }

  const named = index.findModulesByName(word.name, currentPath);
  if (named.length) return chrPreview.appendHover(formatModuleHover(named[0]), lineText, word);

  const here = index.moduleAt(currentPath);
  if (here) {
    const local = findMember(here, word.name, { includePrivate: true });
    if (local.length) return chrPreview.appendHover(formatMemberHover(here, local), lineText, word);
    const decl = localDecl(here, lineNumber, word.name);
    if (decl) {
      return chrPreview.appendHover(
        formatMemberHover({ name: decl.hostName, members: [decl.mem] }, [decl.mem]),
        lineText,
        word
      );
    }
  }

  const g = builtins.findGlobal(word.name);
  if (g && g.member) {
    return chrPreview.appendHover(
      formatMemberHover({ name: g.owner || 'VBA', members: [g.member] }, [g.member]),
      lineText,
      word
    );
  }
  if (g && g.typeName) {
    const t = builtins.findType(g.typeName);
    if (t) return chrPreview.appendHover(formatModuleHover({ name: t.name, kind: t.kind, subtree: t.library }), lineText, word);
  }
  const t = builtins.findType(word.name);
  if (t) return chrPreview.appendHover(formatModuleHover({ name: t.name, kind: t.kind, subtree: t.library }), lineText, word);

  const defs = definitionsAt(index, currentPath, lineText, lineNumber, character, documentText);
  if (defs.length) return chrPreview.appendHover(defs[0].name, lineText, word);
  return chrPreview.appendHover(null, lineText, word);
}

module.exports = {
  wordAt,
  qualifierBefore,
  memberTrigger,
  resolveTypeName,
  moduleForType,
  findMember,
  uniqueMembers,
  definitionsAt,
  completionsAt,
  hoverAt,
  formatMemberHover
};
