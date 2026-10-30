'use strict';

const path = require('path');

const IDENT = '[A-Za-z_][A-Za-z0-9_]*';
const MEMBER_RE = new RegExp(
  `^\\s*(?:(?<access>Public|Friend|Private)\\s+)?(?:Static\\s+)?(?<kind>Property\\s+(?<prop>Get|Let|Set)|Function|Sub)\\s+(?<name>${IDENT})(?<rest>.*)$`,
  'i'
);
const ATTR_NAME_RE = /Attribute\s+VB_Name\s*=\s*"([^"]+)"/i;
const ATTR_PREDECLARED_RE = /Attribute\s+VB_PredeclaredId\s*=\s*(True|False)/i;
const AS_TYPE_RE = new RegExp(`\\bAs\\s+(?:New\\s+)?(${IDENT})\\s*$`, 'i');
const DIM_RE = new RegExp(
  `^\\s*(?:Dim|Private|Public|Static)\\s+(${IDENT})(?:\\s*,\\s*${IDENT})*(?:\\s+As\\s+(?:New\\s+)?(${IDENT}))?`,
  'i'
);
const CONST_RE = new RegExp(
  `^\\s*(?:(?<access>Public|Friend|Private)\\s+)?Const\\s+(?<name>${IDENT})(?:\\s+As\\s+(?<type>${IDENT}))?\\s*=\\s*(?<value>.+)$`,
  'i'
);
const FIELD_RE = new RegExp(
  `^\\s*(?:(?<access>Dim|Private|Public|Static)\\s+)(?:WithEvents\\s+)?(?<name>${IDENT})(?:\\(\\))?\\s*(?:As\\s+(?:New\\s+)?(?<type>${IDENT}))?\\s*$`,
  'i'
);
const SET_NEW_RE = new RegExp(`^\\s*Set\\s+(${IDENT})\\s*=\\s*New\\s+(${IDENT})`, 'i');
const WITH_RE = new RegExp(`^\\s*With\\s+(${IDENT})(?:\\s*\\.|\\s|$)`, 'i');

function titleAccess(access) {
  const a = access || '';
  if (!a) return '';
  return a.charAt(0).toUpperCase() + a.slice(1).toLowerCase();
}

function makeDecl(name, kind, access, signature, line, type, doc) {
  const acc = titleAccess(access) || (kind === 'Const' ? 'Public' : 'Private');
  return {
    name,
    kind,
    access: acc,
    isPrivate: !/^public$/i.test(acc),
    signature,
    line,
    returnType: type || '',
    declParams: [],
    doc: doc || { summary: '', params: [], returns: '' },
    params: []
  };
}

function kindFromExt(fsPath) {
  const ext = path.extname(fsPath).toLowerCase();
  if (ext === '.cls') return 'class';
  if (ext === '.bas') return 'standard';
  if (ext === '.frm') return 'form';
  if (ext === '.dcm') return 'document';
  return 'standard';
}

function subtreeOf(fsPath) {
  return path.basename(path.dirname(path.resolve(fsPath)));
}

function defaultName(fsPath) {
  return path.basename(fsPath, path.extname(fsPath));
}

function controlTypeName(progId) {
  const match = String(progId || '').match(/Forms\.([A-Za-z]+)/i);
  if (match) return match[1];
  return '';
}

function splitLines(text) {
  return text.split(/\r?\n/);
}

function logicalLines(text) {
  const raw = splitLines(text);
  const out = [];
  let buf = '';
  let start = 0;
  for (let i = 0; i < raw.length; i++) {
    const line = raw[i];
    if (!buf) start = i;
    const trimmedEnd = line.replace(/[\t ]+$/, '');
    if (/[ \t]_$/.test(trimmedEnd)) {
      buf += trimmedEnd.slice(0, -1);
      continue;
    }
    out.push({ text: buf + line, line: start, endLine: i });
    buf = '';
  }
  if (buf) out.push({ text: buf, line: start, endLine: raw.length - 1 });
  return out;
}

function stripTrailingComment(line) {
  let inStr = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inStr && line[i + 1] === '"') {
        i += 1;
        continue;
      }
      inStr = !inStr;
    } else if (ch === "'" && !inStr) {
      return line.slice(0, i).replace(/[\t ]+$/, '');
    }
  }
  return line;
}

function splitParamList(rest) {
  const text = rest || '';
  let depth = 0;
  let close = -1;
  for (let i = 0; i < text.length; i++) {
    if (text[i] === '(') depth += 1;
    else if (text[i] === ')') {
      depth -= 1;
      if (depth === 0) {
        close = i;
        break;
      }
    }
  }
  if (close < 0) return { inner: '', after: text };
  return { inner: text.slice(1, close), after: text.slice(close + 1) };
}

function splitTopLevel(list) {
  const parts = [];
  let buf = '';
  let depth = 0;
  for (let i = 0; i < list.length; i++) {
    const ch = list[i];
    if (ch === '(') depth += 1;
    if (ch === ')') depth -= 1;
    if (ch === ',' && depth === 0) {
      if (buf.trim()) parts.push(buf.trim());
      buf = '';
    } else {
      buf += ch;
    }
  }
  if (buf.trim()) parts.push(buf.trim());
  return parts;
}

function parseOneParam(part) {
  const optional = /\bOptional\b/i.test(part);
  const paramArray = /\bParamArray\b/i.test(part);
  let rest = part.replace(/^\s*(?:(?:Optional|ParamArray|ByVal|ByRef)\s+)+/i, '').trim();
  let defaultValue = '';
  const eq = rest.match(/^(.*?)\s*=\s*(.+)$/);
  if (eq) {
    rest = eq[1].trim();
    defaultValue = eq[2].trim();
  }
  let type = '';
  const asM = rest.match(/^(.*?)\s+As\s+(?:New\s+)?(.+)$/i);
  if (asM) {
    rest = asM[1].trim();
    type = asM[2].trim();
  }
  const name = rest.replace(/\(\s*\)$/, '').trim();
  return { name, type, optional, paramArray, defaultValue, text: '' };
}

function parseDeclParams(rest) {
  const { inner, after } = splitParamList(rest);
  const params = inner ? splitTopLevel(inner).map(parseOneParam).filter((p) => p.name) : [];
  const afterClean = stripTrailingComment(after);
  const ret = afterClean.match(/^\s*As\s+(?:New\s+)?([A-Za-z_]\w*(?:\(\))?)/i);
  return { params, returnType: ret ? ret[1] : '' };
}

function returnTypeOf(signature) {
  const cleaned = stripTrailingComment(signature).replace(/[\t ]+$/, '');
  const nameM = cleaned.match(MEMBER_RE);
  if (nameM && nameM.groups) {
    const parsed = parseDeclParams(nameM.groups.rest);
    if (parsed.returnType) return parsed.returnType;
  }
  const m = cleaned.match(AS_TYPE_RE);
  return m ? m[1] : '';
}

function isCommentLine(line) {
  return /^\s*'/.test(line);
}

function commentText(line) {
  return line.replace(/^\s*'/, '').replace(/^\s/, '');
}

function parseDocBlock(rawLines, skip, procLine) {
  const comments = [];
  for (let i = procLine - 1; i >= 0; i--) {
    if (skip[i]) continue;
    const line = rawLines[i];
    if (/^\s*$/.test(line)) break;
    if (isCommentLine(line)) {
      comments.push(commentText(line));
      continue;
    }
    break;
  }
  comments.reverse();
  const doc = { summary: '', params: [], returns: '' };
  const summary = [];
  for (const text of comments) {
    const param = text.match(/^@param\s+([A-Za-z_]\w*)\s*(.*)$/i);
    if (param) {
      doc.params.push({ name: param[1], text: param[2].trim() });
      continue;
    }
    const ret = text.match(/^@returns?\s+(.*)$/i);
    if (ret) {
      doc.returns = ret[1].trim();
      continue;
    }
    if (text) summary.push(text);
  }
  doc.summary = summary.join(' ').replace(/\s+/g, ' ').trim();
  return doc;
}

function mergeParams(declParams, doc) {
  const byName = new Map();
  for (const p of doc.params || []) byName.set(p.name.toLowerCase(), p.text);
  const names = new Set();
  const out = [];
  for (const p of declParams || []) {
    names.add(p.name.toLowerCase());
    out.push({
      name: p.name,
      type: p.type,
      optional: p.optional,
      paramArray: p.paramArray,
      defaultValue: p.defaultValue,
      text: byName.get(p.name.toLowerCase()) || ''
    });
  }
  for (const p of doc.params || []) {
    if (!names.has(p.name.toLowerCase())) {
      out.push({ name: p.name, type: '', optional: false, paramArray: false, defaultValue: '', text: p.text });
    }
  }
  return out;
}

function skipDesigner(rawLines) {
  const skip = new Array(rawLines.length).fill(false);
  let depth = 0;
  for (let i = 0; i < rawLines.length; i++) {
    const t = rawLines[i].replace(/^\s+/, '');
    if (/^Begin\b/i.test(t)) {
      depth += 1;
      skip[i] = true;
      continue;
    }
    if (depth > 0) {
      skip[i] = true;
      if (/^End\s*$/i.test(t)) depth -= 1;
    }
  }
  return skip;
}

function parseVarsFromLine(line, vars) {
  const cleaned = stripTrailingComment(line);
  const setNew = cleaned.match(SET_NEW_RE);
  if (setNew) {
    vars.set(setNew[1].toLowerCase(), setNew[2]);
    return;
  }
  const dim = cleaned.match(DIM_RE);
  if (dim && dim[2]) vars.set(dim[1].toLowerCase(), dim[2]);
}

function parseModule(fsPath, text) {
  const raw = splitLines(text);
  const skip = skipDesigner(raw);
  const kind = kindFromExt(fsPath);
  const info = {
    fsPath,
    name: defaultName(fsPath),
    kind,
    predeclaredId: kind === 'form',
    subtree: subtreeOf(fsPath),
    members: [],
    globals: [],
    moduleVars: new Map(),
    procedures: []
  };

  for (let i = 0; i < raw.length; i++) {
    if (skip[i]) continue;
    const line = raw[i];
    const nameM = line.match(ATTR_NAME_RE);
    if (nameM) info.name = nameM[1];
    const preM = line.match(ATTR_PREDECLARED_RE);
    if (preM) info.predeclaredId = /^true$/i.test(preM[1]);
  }

  const logical = logicalLines(text);
  let currentProc = null;

  for (const row of logical) {
    if (skip[row.line]) continue;
    const cleaned = stripTrailingComment(row.text);
    const trimmed = cleaned.replace(/^\s+/, '');
    if (!trimmed || trimmed.startsWith("'")) continue;
    if (/^Attribute\s+/i.test(trimmed)) continue;
    if (/^VERSION\b/i.test(trimmed)) continue;
    if (/^Option\s+/i.test(trimmed)) continue;

    const mem = trimmed.match(MEMBER_RE);
    if (mem && mem.groups) {
      const access = (mem.groups.access || 'Public').replace(/^./, (c) => c.toUpperCase());
      const isPrivate = /^private$/i.test(access);
      const kindName = mem.groups.prop
        ? 'Property ' + mem.groups.prop.charAt(0).toUpperCase() + mem.groups.prop.slice(1).toLowerCase()
        : mem.groups.kind.charAt(0).toUpperCase() + mem.groups.kind.slice(1).toLowerCase();
      const signature = cleaned.replace(/^\s+/, '').replace(/\s+/g, ' ');
      const decl = parseDeclParams(mem.groups.rest);
      const doc = parseDocBlock(raw, skip, row.line);
      const member = {
        name: mem.groups.name,
        kind: kindName,
        access,
        isPrivate,
        signature,
        line: row.line,
        returnType: decl.returnType || returnTypeOf(cleaned),
        declParams: decl.params,
        doc,
        params: mergeParams(decl.params, doc)
      };
      info.members.push(member);

      if (currentProc) currentProc.endLine = row.line - 1;
      currentProc = {
        name: member.name,
        line: row.line,
        endLine: raw.length - 1,
        vars: new Map(),
        locals: new Map(),
        params: member.params || []
      };
      for (const p of currentProc.params) {
        if (p.name && p.type) currentProc.vars.set(p.name.toLowerCase(), p.type);
      }
      info.procedures.push(currentProc);

      if (kind === 'standard' && !isPrivate && /^property get$/i.test(kindName) && member.returnType) {
        info.globals.push({
          name: member.name,
          typeName: member.returnType,
          line: row.line
        });
      }
      continue;
    }

    if (/^End\s+(Sub|Function|Property)\b/i.test(trimmed)) {
      if (currentProc) currentProc.endLine = row.endLine;
      currentProc = null;
      continue;
    }

    const cst = trimmed.match(CONST_RE);
    if (cst && cst.groups) {
      const access = cst.groups.access || (currentProc ? 'Private' : 'Public');
      const signature = cleaned.replace(/^\s+/, '').replace(/\s+/g, ' ');
      const doc = parseDocBlock(raw, skip, row.line);
      const decl = makeDecl(cst.groups.name, 'Const', access, signature, row.line, cst.groups.type, doc);
      if (currentProc) {
        currentProc.locals.set(decl.name.toLowerCase(), decl);
        if (decl.returnType) currentProc.vars.set(decl.name.toLowerCase(), decl.returnType);
      } else {
        info.members.push(decl);
        if (decl.returnType) info.moduleVars.set(decl.name.toLowerCase(), decl.returnType);
      }
      continue;
    }

    const field = trimmed.match(FIELD_RE);
    if (field && field.groups) {
      const access = field.groups.access;
      const signature = cleaned.replace(/^\s+/, '').replace(/\s+/g, ' ');
      const doc = parseDocBlock(raw, skip, row.line);
      const decl = makeDecl(field.groups.name, 'Variable', access, signature, row.line, field.groups.type, doc);
      if (currentProc) {
        currentProc.locals.set(decl.name.toLowerCase(), decl);
        if (decl.returnType) currentProc.vars.set(decl.name.toLowerCase(), decl.returnType);
      } else {
        info.members.push(decl);
        if (decl.returnType) info.moduleVars.set(decl.name.toLowerCase(), decl.returnType);
      }
      continue;
    }

    if (currentProc) parseVarsFromLine(cleaned, currentProc.vars);
    else parseVarsFromLine(cleaned, info.moduleVars);
  }

  if (info.predeclaredId && (kind === 'class' || kind === 'form')) {
    const already = info.globals.some((g) => g.name.toLowerCase() === info.name.toLowerCase());
    if (!already) {
      info.globals.push({ name: info.name, typeName: info.name, line: 0 });
    }
  }

  return info;
}

class VbaIndex {
  constructor() {
    this.byPath = new Map();
  }

  ingest(fsPath, text) {
    const resolved = path.resolve(fsPath);
    this.byPath.set(resolved, parseModule(resolved, text));
  }

  remove(fsPath) {
    this.byPath.delete(path.resolve(fsPath));
  }

  all() {
    return [...this.byPath.values()];
  }

  moduleAt(fsPath) {
    return this.byPath.get(path.resolve(fsPath));
  }

  preferred(modules, currentPath) {
    const here = currentPath ? path.resolve(currentPath) : '';
    const dir = here ? path.dirname(here).toLowerCase() : '';
    const rank = (mod) => {
      if (here && path.resolve(mod.fsPath) === here) return 0;
      if (dir && path.dirname(path.resolve(mod.fsPath)).toLowerCase() === dir) return 1;
      return 2;
    };
    return modules.slice().sort((a, b) => rank(a) - rank(b));
  }

  setControls(fsPath, controls) {
    const mod = this.moduleAt(fsPath);
    if (!mod) return;
    mod.members = mod.members.filter((m) => m.kind !== 'Control');
    mod.controls = [];
    for (const control of controls || []) {
      const name = control && control.name;
      if (!name) continue;
      const typeName = control.type || controlTypeName(control.progId);
      const member = {
        name,
        kind: 'Control',
        access: 'Public',
        isPrivate: false,
        signature: typeName ? name + ' As ' + typeName : name,
        line: 0,
        returnType: typeName || '',
        declParams: [],
        doc: { summary: '', params: [], returns: '' },
        params: []
      };
      mod.controls.push({ name, progId: control.progId || '', type: typeName });
      mod.members.push(member);
    }
  }

  findModulesByName(name, currentPath) {
    const lower = name.toLowerCase();
    const matches = this.all().filter((m) => m.name.toLowerCase() === lower);
    return this.preferred(matches, currentPath);
  }

  findGlobals(name, currentPath) {
    const lower = name.toLowerCase();
    const hits = [];
    for (const mod of this.all()) {
      for (const g of mod.globals) {
        if (g.name.toLowerCase() === lower) hits.push({ module: mod, global: g });
      }
    }
    const mods = this.preferred(hits.map((h) => h.module), currentPath);
    return hits.sort((a, b) => mods.indexOf(a.module) - mods.indexOf(b.module));
  }
}

function enclosingWith(text, lineNumber) {
  const raw = splitLines(text);
  let depth = 0;
  for (let i = lineNumber; i >= 0; i--) {
    const cleaned = stripTrailingComment(raw[i]);
    if (/^\s*End\s+With\b/i.test(cleaned)) {
      depth += 1;
      continue;
    }
    const m = cleaned.match(WITH_RE);
    if (m) {
      if (depth === 0) return m[1];
      depth -= 1;
    }
  }
  return '';
}

function enclosingProc(mod, lineNumber) {
  if (!mod || !mod.procedures) return null;
  for (const proc of mod.procedures) {
    if (lineNumber >= proc.line && lineNumber <= proc.endLine) return proc;
  }
  return null;
}

function varsForLine(mod, lineNumber) {
  const vars = new Map(mod.moduleVars);
  for (const proc of mod.procedures) {
    if (lineNumber >= proc.line && lineNumber <= proc.endLine) {
      for (const [k, v] of proc.vars) vars.set(k, v);
      break;
    }
  }
  return vars;
}

module.exports = {
  IDENT,
  VbaIndex,
  parseModule,
  subtreeOf,
  logicalLines,
  enclosingWith,
  enclosingProc,
  varsForLine,
  stripTrailingComment,
  parseDeclParams,
  parseDocBlock,
  mergeParams
};
