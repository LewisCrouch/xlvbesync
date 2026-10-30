'use strict';

const fs = require('fs');
const path = require('path');
const { VbaIndex } = require('./src/index');
const { definitionsAt, completionsAt, hoverAt } = require('./src/resolve');
const chrPreview = require('./src/chrPreview');
const formEdit = require('./src/formEdit');

function walk(dir, acc) {
  for (const name of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, name.name);
    if (name.isDirectory()) walk(p, acc);
    else if (/\.(cls|bas|frm|dcm)$/i.test(name.name)) acc.push(p);
  }
  return acc;
}

function fail(msg) {
  console.error('FAIL ' + msg);
  process.exitCode = 1;
}

function ok(msg) {
  console.log('ok   ' + msg);
}

const root = path.join(__dirname, 'fixtures');
const index = new VbaIndex();
for (const file of walk(root, [])) {
  index.ingest(file, fs.readFileSync(file, 'utf8'));
}

const csvPath = path.join(root, 'book', 'CSV.cls');
const mainPath = path.join(root, 'book', 'Main.bas');
const twBook = path.join(root, 'book', 'ThisWorkbook.dcm');
const twAddin = path.join(root, 'addin', 'ThisWorkbook.dcm');
const utilPath = path.join(root, 'book', 'Util.cls');
const formPath = path.join(root, 'book', 'frmDemo.frm');

const csv = index.moduleAt(csvPath);
if (!csv || csv.name !== 'CSV') fail('CSV module name'); else ok('CSV module');
if (!csv.members.some((m) => m.name === 'GetColumn')) fail('CSV.GetColumn'); else ok('CSV.GetColumn indexed');

const factory = index.moduleAt(path.join(root, 'book', 'Factory.bas'));
if (!factory.globals.some((g) => g.name === 'CSV' && g.typeName === 'CSV')) fail('factory CSV global');
else ok('Factory CSV global');

const util = index.moduleAt(utilPath);
if (!util.predeclaredId) fail('Util PredeclaredId'); else ok('Util PredeclaredId');

const form = index.moduleAt(formPath);
if (!form || form.kind !== 'form') fail('frmDemo kind'); else ok('frmDemo indexed as a form');
if (form.members.some((m) => m.name === 'Caption' || m.name === 'Begin')) fail('designer block leaked into symbols');
else ok('designer block is not indexed');
if (!form.members.some((m) => m.name === 'btnGo_Click')) fail('btnGo_Click'); else ok('form code-behind indexed');

let defs = definitionsAt(index, mainPath, '  CSV.GetColumn', 0, 4, '  CSV.GetColumn');
if (!defs.length || path.basename(defs[0].fsPath) !== 'CSV.cls') fail('click CSV from Main');
else ok('Ctrl+click CSV from Main → CSV.cls');

defs = definitionsAt(index, mainPath, '  CSV.GetColumn', 0, 10, '  CSV.GetColumn');
if (!defs.length || defs[0].name !== 'GetColumn') fail('click GetColumn from Main');
else ok('CSV.GetColumn member');

const csvMembers = completionsAt(index, mainPath, '  CSV.', 0, 6, '  CSV.');
if (!csvMembers.some((c) => c.name === 'GetColumn')) fail('complete CSV.');
else ok('completion after CSV. includes GetColumn');

const utilMembers = completionsAt(index, mainPath, '  Util.', 0, 7, '  Util.');
if (!utilMembers.some((c) => c.name === 'StartsWith')) fail('complete Util.');
else ok('Util. includes StartsWith');

defs = definitionsAt(index, mainPath, '  ThisWorkbook', 0, 4, '  ThisWorkbook');
if (!defs.length || path.normalize(defs[0].fsPath) !== path.normalize(twBook)) fail('ThisWorkbook prefers book');
else ok('ThisWorkbook from Main → book');

defs = definitionsAt(index, path.join(root, 'addin', 'ThisWorkbook.dcm'), '  ThisWorkbook', 0, 4, '  ThisWorkbook');
if (!defs.length || path.normalize(defs[0].fsPath) !== path.normalize(twAddin)) fail('ThisWorkbook prefers addin');
else ok('ThisWorkbook from add-in → addin');

const hover = hoverAt(index, mainPath, '  CSV.GetColumn', 0, 10, '  CSV.GetColumn');
if (!hover || !/GetColumn/.test(hover)) fail('hover GetColumn');
else ok('hover GetColumn signature');

const joinHover = hoverAt(index, mainPath, '  FS.JoinPaths', 0, 8, '  FS.JoinPaths');
if (!joinHover || !/Join path fragments/.test(joinHover) || !/combined path/.test(joinHover)) fail('hover FS.JoinPaths docs');
else ok('hover FS.JoinPaths includes summary and returns');

const lenLine = '  n = Len(s)';
const lenHover = hoverAt(index, mainPath, lenLine, 0, lenLine.indexOf('Len') + 1, lenLine);
if (!lenHover || !/\bLen\b/.test(lenHover) || !/characters/.test(lenHover)) fail('hover Len');
else ok('hover Len includes VBA summary');

if (!chrPreview.describe(13) || chrPreview.describe(13).glyph !== '⏎') fail('Chr 13 glyph');
else ok('Chr(13) previews as return');
if (!chrPreview.describe(32) || chrPreview.describe(32).glyph !== '␣') fail('Chr 32 glyph');
else ok('Chr(32) previews as space');

const chrCalls = chrPreview.findCalls('  s = Chr$(13) & Chr$(10) \' Chr$(32)');
if (chrCalls.length !== 2 || chrCalls[0].glyph !== '⏎' || chrCalls[1].glyph !== '↵') fail('find Chr$ calls');
else ok('finds Chr$ return and line feed, skips comments');

const appLine = '  Application.EnableEvents = False';
const appHover = hoverAt(index, mainPath, appLine, 0, appLine.indexOf('EnableEvents') + 1, appLine);
if (!appHover || !/EnableEvents/.test(appHover)) fail('hover Application.EnableEvents');
else ok('hover Application.EnableEvents');

const appMembers = completionsAt(index, mainPath, '  Application.', 0, 14, '  Application.');
if (!appMembers.some((c) => c.name === 'EnableEvents') || !appMembers.some((c) => c.name === 'Workbooks')) fail('complete Application.');
else ok('completion after Application. includes EnableEvents and Workbooks');

const sheetsLine = '  ThisWorkbook.Sheets';
const sheetsHover = hoverAt(index, mainPath, sheetsLine, 0, sheetsLine.indexOf('Sheets') + 1, sheetsLine);
if (!sheetsHover || !/Sheets/.test(sheetsHover)) fail('hover ThisWorkbook.Sheets');
else ok('hover ThisWorkbook.Sheets uses Excel Workbook');

const fsPath = path.join(root, 'book', 'FS.cls');
const fsText = fs.readFileSync(fsPath, 'utf8');
const fsLines = fsText.split(/\r?\n/);
const constDeclLine = fsLines.findIndex((l) => /Private Const cMaxFilePathLength/.test(l));
const constUseLine = fsLines.findIndex((l) => /MaxFilePathLength = cMaxFilePathLength/.test(l));
const paramUseLine = fsLines.findIndex((l) => /IsWithinFilePathLimit = Len\(filePath\)/.test(l));
const paramDeclLine = fsLines.findIndex((l) => /Function IsWithinFilePathLimit\(filePath/.test(l));

if (constDeclLine < 0 || constUseLine < 0) fail('FS const lines');
else {
  const constHover = hoverAt(
    index, fsPath, fsLines[constUseLine], constUseLine,
    fsLines[constUseLine].indexOf('cMaxFilePathLength') + 1, fsText
  );
  if (!constHover || !/cMaxFilePathLength/.test(constHover) || !/218/.test(constHover)) fail('hover cMaxFilePathLength');
  else ok('hover cMaxFilePathLength shows Const');
}

if (paramDeclLine < 0 || paramUseLine < 0) fail('FS filePath lines');
else {
  const paramHover = hoverAt(
    index, fsPath, fsLines[paramUseLine], paramUseLine,
    fsLines[paramUseLine].indexOf('filePath') + 1, fsText
  );
  if (!paramHover || !/filePath/.test(paramHover) || !/String/.test(paramHover) || !/Path to test/.test(paramHover)) fail('hover filePath');
  else ok('hover filePath shows parameter');
}

const extra = '\nPublic Sub Foo()\n  Dim cache As CSV\n  cache.GetColumn\n  Dim ws As Worksheet\n  ws.Name\nEnd Sub\n';
const mainPlus = fs.readFileSync(mainPath, 'utf8') + extra;
index.ingest(mainPath, mainPlus);
const extraLines = mainPlus.split(/\r?\n/);
const useLine = extraLines.findIndex((l) => /^\s*cache\./.test(l));
const afterDim = completionsAt(index, mainPath, extraLines[useLine], useLine, extraLines[useLine].indexOf('.') + 1, mainPlus);
if (!afterDim.some((c) => c.name === 'GetColumn')) fail('Dim cache As CSV then cache.');
else ok('Dim As CSV then cache. includes GetColumn');

const wsLine = extraLines.findIndex((l) => /^\s*ws\.Name\s*$/.test(l));
const afterWs = completionsAt(index, mainPath, extraLines[wsLine], wsLine, extraLines[wsLine].indexOf('.') + 1, mainPlus);
if (!afterWs.some((c) => c.name === 'Name') || !afterWs.some((c) => c.name === 'Cells')) fail('Dim ws As Worksheet then ws.');
else ok('Dim As Worksheet then ws. includes Name and Cells');

index.ingest(formPath, fs.readFileSync(formPath, 'utf8'));
index.setControls(formPath, [{ name: 'btnGo', progId: 'Forms.CommandButton.1' }]);
const btnLine = '  btnGo.';
const btnMembers = completionsAt(index, formPath, btnLine, 0, btnLine.indexOf('.') + 1, btnLine);
if (!btnMembers.some((c) => c.name === 'Caption') || !btnMembers.some((c) => c.name === 'SetFocus')) fail('btnGo. members');
else ok('btnGo. includes Caption and SetFocus');

if (formEdit.snap(13, 6, true) !== 12) fail('snap 13 to 6');
else ok('snap rounds to the grid');
if (formEdit.snap(13, 6, false) !== 13) fail('snap disabled');
else ok('snap leaves the value when SnapToGrid is off');
if (formEdit.nextName('CommandButton', ['commandbutton1']) !== 'CommandButton2') fail('next control name');
else ok('next control name skips names in use');
if (formEdit.defaultEvent('Forms.TextBox.1') !== 'Change') fail('textbox default event');
else ok('TextBox default event is Change');
if (formEdit.defaultEvent('Forms.CommandButton.1') !== 'Click') fail('button default event');
else ok('CommandButton default event is Click');
if (!/Private Sub btnGo_Click\(/.test(formEdit.eventStub('btnGo', 'Click'))) fail('event stub');
else ok('event stub');

if (process.exitCode) {
  console.error('some checks failed');
} else {
  console.log('all checks passed (' + index.all().length + ' modules)');
}
