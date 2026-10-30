'use strict';

const path = require('path');
const vscode = require('vscode');
const { VbaIndex } = require('./index');
const { definitionsAt, completionsAt, hoverAt } = require('./resolve');
const chrPreview = require('./chrPreview');
const { registerDesigner } = require('./designer');

const SELECTORS = [
  { language: 'vba', scheme: 'file' },
  { language: 'vb', scheme: 'file' }
];
const SOURCE_EXTS = ['cls', 'bas', 'frm', 'dcm'];

function completionKind(kind) {
  if (kind.startsWith('Property')) return vscode.CompletionItemKind.Property;
  if (kind === 'Function') return vscode.CompletionItemKind.Function;
  if (kind === 'Sub') return vscode.CompletionItemKind.Method;
  if (kind === 'Const') return vscode.CompletionItemKind.EnumMember;
  if (kind === 'Parameter' || kind === 'Variable') return vscode.CompletionItemKind.Variable;
  if (kind === 'Class') return vscode.CompletionItemKind.Class;
  if (kind === 'form') return vscode.CompletionItemKind.Class;
  if (kind === 'Control') return vscode.CompletionItemKind.Field;
  if (kind === 'Module') return vscode.CompletionItemKind.Module;
  return vscode.CompletionItemKind.Text;
}

function locationOf(def) {
  return new vscode.Location(vscode.Uri.file(def.fsPath), new vscode.Position(def.line, 0));
}

function isVbaDocument(document) {
  if (!document) return false;
  if (document.languageId === 'vba' || document.languageId === 'vb') return true;
  return /\.(cls|bas|frm|dcm)$/i.test(document.fileName || '');
}

async function srcDirFor(folder) {
  try {
    const uri = vscode.Uri.joinPath(folder.uri, 'xlvbesync.json');
    const bytes = await vscode.workspace.fs.readFile(uri);
    const cfg = JSON.parse(Buffer.from(bytes).toString('utf8'));
    if (typeof cfg.src_dir === 'string' && cfg.src_dir.trim()) {
      return cfg.src_dir.trim().replace(/\\/g, '/').replace(/^\.\//, '').replace(/\/$/, '');
    }
  } catch (_) {
    /* default */
  }
  return 'src';
}

function activate(context) {
  const index = new VbaIndex();
  const controlCache = new Map();

  const ingest = (fsPath, text) => {
    index.ingest(fsPath, text);
    const cached = controlCache.get(path.resolve(fsPath));
    if (cached) index.setControls(fsPath, cached);
  };

  const rememberControls = (fsPath, controls) => {
    const key = path.resolve(fsPath);
    controlCache.set(key, controls);
    if (index.moduleAt(fsPath)) index.setControls(fsPath, controls);
    context.workspaceState.update('xlvbesync.formControls', [...controlCache.entries()]);
  };

  const saved = context.workspaceState.get('xlvbesync.formControls');
  if (Array.isArray(saved)) {
    for (const entry of saved) {
      if (Array.isArray(entry) && entry.length === 2) controlCache.set(entry[0], entry[1]);
    }
  }

  async function scanWorkspace() {
    const folders = vscode.workspace.workspaceFolders || [];
    for (const folder of folders) {
      const srcDir = await srcDirFor(folder);
      const patterns = SOURCE_EXTS.map((ext) => new vscode.RelativePattern(folder, `${srcDir}/**/*.${ext}`));
      const found = await Promise.all(patterns.map((p) => vscode.workspace.findFiles(p, '**/node_modules/**')));
      await Promise.all(found.flat().map(async (uri) => {
        const bytes = await vscode.workspace.fs.readFile(uri);
        ingest(uri.fsPath, Buffer.from(bytes).toString('utf8'));
      }));
    }
  }

  let ready = scanWorkspace();

  const watchers = [];
  const refresh = async (uri) => {
    try {
      const bytes = await vscode.workspace.fs.readFile(uri);
      ingest(uri.fsPath, Buffer.from(bytes).toString('utf8'));
    } catch (_) {
      index.remove(uri.fsPath);
    }
  };

  for (const folder of vscode.workspace.workspaceFolders || []) {
    const watcher = vscode.workspace.createFileSystemWatcher(
      new vscode.RelativePattern(folder, '**/*.{cls,bas,frm,dcm}')
    );
    watcher.onDidCreate(refresh);
    watcher.onDidChange(refresh);
    watcher.onDidDelete((uri) => index.remove(uri.fsPath));
    watchers.push(watcher);
  }

  const overrideFromEditor = (document) => {
    if (isVbaDocument(document)) ingest(document.fileName, document.getText());
  };

  const chrBadge = vscode.window.createTextEditorDecorationType({
    after: {
      margin: '0 0 0 0.4em',
      color: new vscode.ThemeColor('editorCodeLens.foreground'),
      backgroundColor: new vscode.ThemeColor('editorInlayHint.background'),
      fontWeight: '600'
    }
  });

  const paintChrBadges = (editor) => {
    if (!editor) return;
    const doc = editor.document;
    if (!isVbaDocument(doc)) {
      editor.setDecorations(chrBadge, []);
      return;
    }
    editor.setDecorations(chrBadge, chrPreview.findCalls(doc.getText()).map((hit) => ({
      range: new vscode.Range(hit.line, hit.end, hit.line, hit.end),
      renderOptions: { after: { contentText: hit.label } },
      hoverMessage: hit.tooltip
    })));
  };

  const paintOpenEditors = () => {
    for (const editor of vscode.window.visibleTextEditors) paintChrBadges(editor);
  };

  const subscriptions = [
    chrBadge,
    ...watchers,
    vscode.workspace.onDidChangeTextDocument((e) => {
      overrideFromEditor(e.document);
      for (const editor of vscode.window.visibleTextEditors) {
        if (editor.document === e.document) paintChrBadges(editor);
      }
    }),
    vscode.window.onDidChangeActiveTextEditor(paintChrBadges),
    vscode.window.onDidChangeVisibleTextEditors(paintOpenEditors)
  ];

  for (const selector of SELECTORS) {
    subscriptions.push(
      vscode.languages.registerDefinitionProvider(selector, {
        async provideDefinition(document, position) {
          await ready;
          overrideFromEditor(document);
          const line = document.lineAt(position.line).text;
          return definitionsAt(
            index,
            document.fileName,
            line,
            position.line,
            position.character,
            document.getText()
          ).map(locationOf);
        }
      }),
      vscode.languages.registerCompletionItemProvider(selector, {
        async provideCompletionItems(document, position) {
          await ready;
          overrideFromEditor(document);
          const line = document.lineAt(position.line).text;
          const items = completionsAt(
            index,
            document.fileName,
            line,
            position.line,
            position.character,
            document.getText()
          );
          return items.map((entry) => {
            const item = new vscode.CompletionItem(entry.name, completionKind(entry.kind));
            item.detail = entry.detail;
            item.documentation = new vscode.MarkdownString(entry.documentation || entry.signature || '');
            item.sortText = entry.name.toLowerCase();
            return item;
          });
        }
      }, '.'),
      vscode.languages.registerHoverProvider(selector, {
        async provideHover(document, position) {
          await ready;
          overrideFromEditor(document);
          const line = document.lineAt(position.line).text;
          const text = hoverAt(
            index,
            document.fileName,
            line,
            position.line,
            position.character,
            document.getText()
          );
          if (!text) return undefined;
          const md = new vscode.MarkdownString(text);
          md.isTrusted = false;
          return new vscode.Hover(md);
        }
      })
    );
  }

  context.subscriptions.push(...subscriptions);
  registerDesigner(context, { index, rememberControls, ingest });
  paintOpenEditors();
}

function deactivate() {}

module.exports = { activate, deactivate };
