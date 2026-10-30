'use strict';

const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

class FormHostClient {
  constructor(workspaceRoot) {
    this.workspaceRoot = workspaceRoot || '';
    this.proc = null;
    this.seq = 1;
    this.pending = new Map();
    this.stderr = '';
  }

  command() {
    const root = this.workspaceRoot;
    const python = root && path.join(root, '.venv', 'Scripts', 'python.exe');
    if (python && fs.existsSync(python)) {
      return { cmd: python, args: ['-m', 'xlvbesync', 'form-host'], cwd: root };
    }
    const exe = root && path.join(root, '.venv', 'Scripts', 'xlvbesync.exe');
    if (exe && fs.existsSync(exe)) {
      return { cmd: exe, args: ['form-host'], cwd: root };
    }
    return { cmd: 'xlvbesync', args: ['form-host'], cwd: root || undefined };
  }

  start() {
    if (this.proc) return;
    const spec = this.command();
    this.proc = spawn(spec.cmd, spec.args, {
      cwd: spec.cwd,
      windowsHide: true,
      stdio: ['pipe', 'pipe', 'pipe'],
      env: process.env
    });
    this.stderr = '';
    this.proc.stderr.setEncoding('utf8');
    this.proc.stderr.on('data', (chunk) => {
      this.stderr = (this.stderr + chunk).slice(-4000);
    });
    const rl = readline.createInterface({ input: this.proc.stdout });
    rl.on('line', (line) => this._onLine(line));
    this.proc.on('exit', (code) => {
      const detail = this.stderr.trim();
      const error = new Error(detail || ('form-host exited (' + code + ')'));
      for (const pending of this.pending.values()) pending.reject(error);
      this.pending.clear();
      this.proc = null;
    });
    this.proc.on('error', (err) => {
      const error = new Error(
        'Could not start xlvbesync form-host. Install the package (pip install -e .) and retry. ' + err.message
      );
      for (const pending of this.pending.values()) pending.reject(error);
      this.pending.clear();
      this.proc = null;
    });
  }

  request(message, timeoutMs) {
    this.start();
    if (!this.proc || !this.proc.stdin.writable) {
      return Promise.reject(new Error('form-host is not running'));
    }
    const id = this.seq++;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error('Timed out waiting for Excel. If a dialog is open in the designer\'s Excel, close it and retry.'));
      }, timeoutMs || 90000);
      this.pending.set(id, { resolve, reject, timer });
      this.proc.stdin.write(JSON.stringify(Object.assign({ id }, message)) + '\n');
    });
  }

  _onLine(line) {
    if (!line || !line.trim()) return;
    let msg;
    try {
      msg = JSON.parse(line);
    } catch (_) {
      return;
    }
    const pending = this.pending.get(msg.id);
    if (!pending) return;
    clearTimeout(pending.timer);
    this.pending.delete(msg.id);
    if (msg.ok) pending.resolve(msg);
    else {
      const error = new Error(msg.error || 'form-host failed');
      error.remedy = msg.remedy || '';
      pending.reject(error);
    }
  }

  dispose() {
    const proc = this.proc;
    if (!proc) return;
    try {
      proc.stdin.write(JSON.stringify({ id: this.seq++, cmd: 'shutdown' }) + '\n');
    } catch (_) {
      /* process already gone */
    }
    setTimeout(() => {
      try { proc.kill(); } catch (_) { /* already exited */ }
    }, 2000);
    this.proc = null;
  }
}

module.exports = { FormHostClient };
