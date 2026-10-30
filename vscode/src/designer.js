'use strict';

const path = require('path');
const vscode = require('vscode');
const { FormHostClient } = require('./formHost');
const formEdit = require('./formEdit');

const VIEW = 'xlvbesync.formDesigner';

class FormDocument {
  constructor(uri) {
    this.uri = uri;
    this.model = null;
    this.session = null;
    this.post = null;
  }

  dispose() {
    this.post = null;
  }
}

class FormEditorProvider {
  constructor(context, hooks) {
    this.context = context;
    this.hooks = hooks;
    this.host = new FormHostClient((vscode.workspace.workspaceFolders || [])[0]?.uri.fsPath || '');
    this.panels = new Map();
    this.active = null;
    this._onDidChange = new vscode.EventEmitter();
    this.onDidChangeCustomDocument = this._onDidChange.event;
  }

  dispose() {
    this.host.dispose();
  }

  async openCustomDocument(uri) {
    return new FormDocument(uri);
  }

  async resolveCustomEditor(document, panel) {
    this.panels.set(document.uri.toString(), panel);
    this.active = document;
    const webview = panel.webview;
    webview.options = {
      enableScripts: true,
      localResourceRoots: [vscode.Uri.file(this.context.extensionPath)]
    };
    let ready = false;
    const queued = [];
    const post = (message) => {
      if (ready) webview.postMessage(message);
      else queued.push(message);
    };
    document.post = post;
    const incoming = webview.onDidReceiveMessage((message) => {
      if (message && message.type === 'ready') {
        ready = true;
        for (const item of queued) webview.postMessage(item);
        queued.length = 0;
        return;
      }
      this.onMessage(document, message);
    });
    panel.onDidChangeViewState((event) => {
      if (event.webviewPanel.active) this.active = document;
    });
    panel.onDidDispose(() => {
      incoming.dispose();
      if (document.session) this.host.request({ cmd: 'close', session: document.session }).catch(() => {});
      document.session = null;
      if (this.active === document) this.active = null;
      this.panels.delete(document.uri.toString());
    });
    webview.html = this.html(webview);
    post({ type: 'status', text: 'Opening the form in Excel…' });
    try {
      const result = await this.host.request({
        cmd: 'open',
        frm: document.uri.fsPath,
        ansi: await this.ansi()
      });
      document.session = result.session;
      document.model = result.model;
      this.remember(document);
      post({ type: 'model', model: result.model });
    } catch (err) {
      post({ type: 'error', message: err.message, remedy: err.remedy || '' });
    }
  }

  async saveCustomDocument(document) {
    await this.persist(document, document.uri);
  }

  async saveCustomDocumentAs(document, destination) {
    await this.persist(document, destination);
  }

  async revertCustomDocument(document) {
    if (document.session) {
      await this.host.request({ cmd: 'close', session: document.session }).catch(() => {});
      document.session = null;
    }
    const result = await this.host.request({ cmd: 'open', frm: document.uri.fsPath, ansi: await this.ansi() });
    document.session = result.session;
    document.model = result.model;
    this.remember(document);
    if (document.post) document.post({ type: 'model', model: result.model });
  }

  async backupCustomDocument(document, context) {
    const payload = Buffer.from(JSON.stringify(document.model || {}), 'utf8');
    await vscode.workspace.fs.writeFile(context.destination, payload);
  }

  async persist(document, destination) {
    if (!document.session || !document.model) {
      throw new Error('The form is not open in Excel yet.');
    }
    const same = destination.toString() === document.uri.toString();
    const result = await this.host.request({
      cmd: 'save',
      session: document.session,
      model: document.model,
      code: await this.codeOf(document.uri),
      frm: same ? undefined : destination.fsPath,
      ansi: await this.ansi()
    });
    this.remember(document);
    if (result.warnings && result.warnings.length) {
      vscode.window.showWarningMessage('Form saved with warnings: ' + result.warnings.slice(0, 4).join('; '));
    }
  }

  remember(document) {
    if (!document.model || !this.hooks.rememberControls) return;
    this.hooks.rememberControls(document.uri.fsPath, formEdit.walkControls(document.model));
  }

  async onMessage(document, message) {
    if (!message || !message.type || message.type === 'ready') return;
    if (message.type === 'edit') {
      const before = message.before;
      const after = message.after;
      document.model = after;
      this._onDidChange.fire({
        document,
        undo: () => {
          document.model = before;
          if (document.post) document.post({ type: 'model', model: before });
        },
        redo: () => {
          document.model = after;
          if (document.post) document.post({ type: 'model', model: after });
        }
      });
      return;
    }
    if (message.type === 'undo') {
      await vscode.commands.executeCommand('undo');
      return;
    }
    if (message.type === 'redo') {
      await vscode.commands.executeCommand('redo');
      return;
    }
    if (message.type === 'viewCode') {
      await viewCode(document.uri, message.name, message.event);
    }
  }

  async newUserForm() {
    const name = await vscode.window.showInputBox({
      title: 'New UserForm',
      prompt: 'Name of the form module',
      value: 'frmForm',
      validateInput(value) {
        if (!/^[A-Za-z][A-Za-z0-9_]{0,30}$/.test(value || '')) {
          return 'Use a letter, then letters, digits or underscore (31 characters max).';
        }
        return null;
      }
    });
    if (!name) return;
    const folder = await this.formDir();
    if (!folder) return;
    const uri = vscode.Uri.joinPath(folder, name + '.frm');
    try {
      await vscode.workspace.fs.stat(uri);
      vscode.window.showErrorMessage(name + '.frm already exists.');
      return;
    } catch (_) {
      /* new file */
    }
    await vscode.workspace.fs.createDirectory(folder);
    try {
      await this.host.request({ cmd: 'create', frm: uri.fsPath, name, ansi: await this.ansi() });
    } catch (err) {
      vscode.window.showErrorMessage(err.message + (err.remedy ? '\n' + err.remedy : ''));
      return;
    }
    await vscode.commands.executeCommand('vscode.openWith', uri, VIEW);
  }

  async viewActiveCode() {
    if (!this.active) return;
    await viewCode(this.active.uri, this.active.model && this.active.model.name, 'Click');
  }

  async formDir() {
    const folders = vscode.workspace.workspaceFolders || [];
    if (!folders.length) {
      vscode.window.showErrorMessage('Open a folder before creating a UserForm.');
      return null;
    }
    const src = await srcDirOf(folders[0]);
    return vscode.Uri.joinPath(folders[0].uri, ...src.split('/').filter(Boolean));
  }

  async ansi() {
    const folders = vscode.workspace.workspaceFolders || [];
    if (!folders.length) return 'cp1252';
    try {
      const raw = await vscode.workspace.fs.readFile(vscode.Uri.joinPath(folders[0].uri, 'xlvbesync.json'));
      const cfg = JSON.parse(Buffer.from(raw).toString('utf8'));
      if (typeof cfg.ansi_encoding === 'string' && cfg.ansi_encoding.trim()) return cfg.ansi_encoding.trim();
    } catch (_) {
      /* default */
    }
    return 'cp1252';
  }

  async codeOf(uri) {
    const open = vscode.workspace.textDocuments.find((doc) => doc.uri.toString() === uri.toString());
    if (open) return open.getText();
    const bytes = await vscode.workspace.fs.readFile(uri);
    return Buffer.from(bytes).toString('utf8');
  }

  html(webview) {
    const nonce = String(Math.random()).slice(2);
    const css = webview.asWebviewUri(vscode.Uri.file(path.join(this.context.extensionPath, 'media', 'designer', 'designer.css')));
    const lib = webview.asWebviewUri(vscode.Uri.file(path.join(this.context.extensionPath, 'src', 'formEdit.js')));
    const script = webview.asWebviewUri(vscode.Uri.file(path.join(this.context.extensionPath, 'media', 'designer', 'designer.js')));
    const csp = webview.cspSource;
    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${csp} data:; style-src ${csp} 'unsafe-inline'; script-src 'nonce-${nonce}';">
<link rel="stylesheet" href="${css}">
</head>
<body>
<div id="app"></div>
<script nonce="${nonce}" src="${lib}"></script>
<script nonce="${nonce}" src="${script}"></script>
</body>
</html>`;
  }
}

async function srcDirOf(folder) {
  try {
    const raw = await vscode.workspace.fs.readFile(vscode.Uri.joinPath(folder.uri, 'xlvbesync.json'));
    const cfg = JSON.parse(Buffer.from(raw).toString('utf8'));
    if (typeof cfg.src_dir === 'string' && cfg.src_dir.trim()) {
      return cfg.src_dir.trim().replace(/\\/g, '/').replace(/^\.\//, '').replace(/\/$/, '');
    }
  } catch (_) {
    /* default */
  }
  return 'src';
}

async function viewCode(uri, name, eventName) {
  await vscode.commands.executeCommand('vscode.openWith', uri, 'default', {
    viewColumn: vscode.ViewColumn.Beside,
    preview: false
  });
  const doc = await vscode.workspace.openTextDocument(uri);
  if (name && eventName) {
    const signature = new RegExp('\\bSub\\s+' + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '_' + eventName + '\\s*\\(', 'i');
    if (!signature.test(doc.getText())) {
      const nl = doc.eol === vscode.EndOfLine.CRLF ? '\r\n' : '\n';
      const stub = formEdit.eventStub(name, eventName).replace(/\r\n/g, nl);
      const lines = doc.getText().split(/\r?\n/);
      let at = lines.length;
      const option = lines.findIndex((line) => /^\s*Option\s+Explicit/i.test(line));
      if (option >= 0) at = option + 1;
      while (at < lines.length && lines[at] === '') at += 1;
      const edit = new vscode.WorkspaceEdit();
      const prefix = at > 0 && lines[at - 1] !== '' ? nl : '';
      edit.insert(uri, new vscode.Position(Math.min(at, lines.length), 0), prefix + stub + (stub.endsWith(nl) ? '' : nl));
      await vscode.workspace.applyEdit(edit);
    }
  }
  const editor = await vscode.window.showTextDocument(doc, {
    viewColumn: vscode.ViewColumn.Beside,
    preview: false
  });
  if (name && eventName) {
    const signature = new RegExp('\\bSub\\s+' + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '_' + eventName + '\\s*\\(', 'i');
    const lines = doc.getText().split(/\r?\n/);
    const line = lines.findIndex((text) => signature.test(text));
    if (line >= 0) {
      editor.selection = new vscode.Selection(line, 0, line, 0);
      editor.revealRange(new vscode.Range(line, 0, Math.min(lines.length, line + 3), 0));
    }
  }
}

function registerDesigner(context, hooks) {
  const provider = new FormEditorProvider(context, hooks);
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider(VIEW, provider, {
      webviewOptions: { retainContextWhenHidden: true },
      supportsMultipleEditorsPerDocument: false
    }),
    vscode.commands.registerCommand('xlvbesync.newUserForm', () => provider.newUserForm()),
    vscode.commands.registerCommand('xlvbesync.viewCode', () => provider.viewActiveCode()),
    { dispose: () => provider.dispose() }
  );
}

module.exports = { registerDesigner };
