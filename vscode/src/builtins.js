'use strict';

const fs = require('fs');
const path = require('path');

const TYPE_ALIASES = {
  activesheet: 'Worksheet',
  activeworkbook: 'Workbook',
  activecell: 'Range',
  selection: 'Range',
  err: 'ErrObject'
};

function normalizeMember(m, summaries) {
  const name = m.name || '';
  const stem = name.replace(/\$$/, '');
  const summary = (summaries[name] || summaries[stem] || m.summary || '').trim();
  return {
    name,
    kind: m.kind || 'Function',
    signature: m.signature || '',
    params: (m.params || []).map((p) => ({
      name: p.name,
      type: p.type || '',
      optional: !!p.optional,
      paramArray: !!p.paramArray,
      defaultValue: p.defaultValue || '',
      text: p.text || ''
    })),
    returnType: m.returnType || '',
    isPrivate: false,
    doc: { summary, returns: '' }
  };
}

function indexMembers(members) {
  const map = new Map();
  for (const m of members) {
    const k = m.name.toLowerCase();
    if (!map.has(k)) map.set(k, []);
    map.get(k).push(m);
  }
  return map;
}

function loadCatalog() {
  const dataPath = path.join(__dirname, '..', 'data', 'typelibs.json');
  const summariesPath = path.join(__dirname, '..', 'data', 'vba-summaries.json');
  const raw = JSON.parse(fs.readFileSync(dataPath, 'utf8'));
  let summaries = {};
  try {
    summaries = JSON.parse(fs.readFileSync(summariesPath, 'utf8'));
  } catch (_) {
    summaries = {};
  }

  const types = new Map();
  for (const ty of Object.values(raw.types || {})) {
    const members = (ty.members || []).map((m) => normalizeMember(m, summaries));
    types.set(ty.name.toLowerCase(), {
      name: ty.name,
      library: ty.library || '',
      kind: ty.kind || 'class',
      members,
      byName: indexMembers(members)
    });
  }

  const globals = new Map();
  for (const g of Object.values(raw.globals || {})) {
    const mem = g.member ? normalizeMember(g.member, summaries) : null;
    let typeName = g.typeName || '';
    const alias = TYPE_ALIASES[(g.name || '').toLowerCase()];
    if (alias) typeName = alias;
    else if (mem && mem.returnType && !typeName) {
      const rt = mem.returnType;
      if (rt && rt !== 'Object' && rt !== 'Variant' && types.has(rt.toLowerCase())) typeName = rt;
    }
    globals.set(g.name.toLowerCase(), {
      name: g.name,
      owner: g.owner || '',
      typeName,
      member: mem
    });
  }

  const formsPath = path.join(__dirname, '..', 'data', 'msforms.json');
  let forms = {};
  try {
    forms = JSON.parse(fs.readFileSync(formsPath, 'utf8'));
  } catch (_) {
    forms = {};
  }
  for (const [typeName, members] of Object.entries(forms)) {
    const key = typeName.toLowerCase();
    if (types.has(key)) continue;
    const normalised = (members || []).map((m) => normalizeMember(m, summaries));
    types.set(key, {
      name: typeName,
      library: 'MSForms',
      kind: 'class',
      members: normalised,
      byName: indexMembers(normalised)
    });
  }

  const intrinsicsPath = path.join(__dirname, '..', 'data', 'vba-intrinsics.json');
  let intrinsics = {};
  try {
    intrinsics = JSON.parse(fs.readFileSync(intrinsicsPath, 'utf8'));
  } catch (_) {
    intrinsics = {};
  }
  for (const rawMem of Object.values(intrinsics)) {
    const name = (rawMem && rawMem.name) || '';
    if (!name || globals.has(name.toLowerCase())) continue;
    const mem = normalizeMember(rawMem, summaries);
    globals.set(name.toLowerCase(), {
      name,
      owner: 'VBA',
      typeName: '',
      member: mem
    });
  }

  return { types, globals };
}

let catalog;

function getCatalog() {
  if (!catalog) catalog = loadCatalog();
  return catalog;
}

function findType(name) {
  if (!name) return null;
  return getCatalog().types.get(name.toLowerCase()) || null;
}

function findGlobal(name) {
  if (!name) return null;
  return getCatalog().globals.get(name.toLowerCase()) || null;
}

function membersOf(typeName) {
  const t = findType(typeName);
  return t ? t.members : [];
}

function asModule(typeRec) {
  if (!typeRec) return null;
  return {
    name: typeRec.name,
    kind: typeRec.kind,
    fsPath: '',
    members: typeRec.members,
    subtree: typeRec.library,
    isBuiltin: true
  };
}

module.exports = {
  getCatalog,
  findType,
  findGlobal,
  membersOf,
  asModule
};
