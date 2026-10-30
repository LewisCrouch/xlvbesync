'use strict';

const vscode = acquireVsCodeApi();
const edit = formEdit;

const PT = 96 / 72;
const TOOLS = [
  ['pointer', 'Select Objects', pointerIcon],
  ['Label', 'Label', labelIcon],
  ['TextBox', 'TextBox', textIcon],
  ['ComboBox', 'ComboBox', comboIcon],
  ['ListBox', 'ListBox', listIcon],
  ['CheckBox', 'CheckBox', checkIcon],
  ['OptionButton', 'OptionButton', optionIcon],
  ['ToggleButton', 'ToggleButton', toggleIcon],
  ['Frame', 'Frame', frameIcon],
  ['CommandButton', 'CommandButton', buttonIcon],
  ['TabStrip', 'TabStrip', tabsIcon],
  ['MultiPage', 'MultiPage', pageIcon],
  ['ScrollBar', 'ScrollBar', scrollIcon],
  ['SpinButton', 'SpinButton', spinIcon],
  ['Image', 'Image', imageIcon]
];

const ENUMS = {
  BorderStyle: [[0, '0 - fmBorderStyleNone'], [1, '1 - fmBorderStyleSingle']],
  SpecialEffect: [[0, '0 - fmSpecialEffectFlat'], [1, '1 - fmSpecialEffectRaised'], [2, '2 - fmSpecialEffectSunken'], [3, '3 - fmSpecialEffectEtched'], [6, '6 - fmSpecialEffectBump']],
  BackStyle: [[0, '0 - fmBackStyleTransparent'], [1, '1 - fmBackStyleOpaque']],
  TextAlign: [[1, '1 - fmTextAlignLeft'], [2, '2 - fmTextAlignCenter'], [3, '3 - fmTextAlignRight']],
  StartUpPosition: [[0, '0 - Manual'], [1, '1 - CenterOwner'], [2, '2 - CenterScreen'], [3, '3 - Windows Default']],
  ScrollBars: [[0, '0 - fmScrollBarsNone'], [1, '1 - fmScrollBarsHorizontal'], [2, '2 - fmScrollBarsVertical'], [3, '3 - fmScrollBarsBoth']],
  PictureSizeMode: [[0, '0 - fmPictureSizeModeClip'], [1, '1 - fmPictureSizeModeStretch'], [3, '3 - fmPictureSizeModeZoom']],
  PictureAlignment: [[0, '0 - fmPictureAlignmentTopLeft'], [1, '1 - fmPictureAlignmentTopRight'], [2, '2 - fmPictureAlignmentCenter'], [3, '3 - fmPictureAlignmentBottomLeft'], [4, '4 - fmPictureAlignmentBottomRight']],
  Cycle: [[0, '0 - fmCycleAllForms'], [2, '2 - fmCycleCurrentForm']],
  Alignment: [[0, '0 - fmAlignmentLeft'], [1, '1 - fmAlignmentRight']],
  MousePointer: [[0, '0 - fmMousePointerDefault'], [1, '1 - fmMousePointerArrow'], [2, '2 - fmMousePointerCross'], [3, '3 - fmMousePointerIBeam'], [11, '11 - fmMousePointerHourGlass'], [14, '14 - fmMousePointerArrowHourGlass'], [15, '15 - fmMousePointerSizeAll'], [99, '99 - fmMousePointerCustom']],
  PicturePosition: [[0, '0 - fmPicturePositionLeftTop'], [1, '1 - fmPicturePositionLeftCenter'], [2, '2 - fmPicturePositionLeftBottom'], [3, '3 - fmPicturePositionRightTop'], [4, '4 - fmPicturePositionRightCenter'], [5, '5 - fmPicturePositionRightBottom'], [6, '6 - fmPicturePositionAboveLeft'], [7, '7 - fmPicturePositionAboveCenter'], [8, '8 - fmPicturePositionAboveRight'], [9, '9 - fmPicturePositionBelowLeft'], [10, '10 - fmPicturePositionBelowCenter'], [11, '11 - fmPicturePositionBelowRight'], [12, '12 - fmPicturePositionCenter']]
};

const CATEGORY = {
  Caption: 'Appearance', Text: 'Appearance', BackColor: 'Appearance', ForeColor: 'Appearance',
  BorderColor: 'Appearance', BorderStyle: 'Appearance', SpecialEffect: 'Appearance', BackStyle: 'Appearance',
  TextAlign: 'Appearance', WordWrap: 'Appearance', Picture: 'Appearance', PicturePosition: 'Appearance',
  PictureSizeMode: 'Appearance', PictureAlignment: 'Appearance', PictureTiling: 'Appearance', Alignment: 'Appearance',
  AutoSize: 'Behavior', Cancel: 'Behavior', Default: 'Behavior', Enabled: 'Behavior', Locked: 'Behavior',
  MultiLine: 'Behavior', PasswordChar: 'Behavior', TabStop: 'Behavior', TabIndex: 'Behavior',
  TakeFocusOnClick: 'Behavior', Visible: 'Behavior', Value: 'Data', MaxLength: 'Data',
  ControlSource: 'Data', RowSource: 'Data', ControlTipText: 'Misc', Accelerator: 'Misc', Tag: 'Misc',
  HelpContextID: 'Misc', MousePointer: 'Misc', GroupName: 'Misc', Zoom: 'Misc',
  Left: 'Position', Top: 'Position', Width: 'Position', Height: 'Position',
  ScrollBars: 'Scrolling', ScrollHeight: 'Scrolling', ScrollWidth: 'Scrolling', ScrollLeft: 'Scrolling', ScrollTop: 'Scrolling',
  KeepScrollBarsVisible: 'Scrolling', Cycle: 'Behavior', ShowModal: 'Behavior', StartUpPosition: 'Behavior',
  GridX: 'Misc', GridY: 'Misc', ShowGrid: 'Misc', SnapToGrid: 'Misc', DrawBuffer: 'Misc', Font: 'Font'
};

const SYSTEM_COLORS = [
  [0x80000000, 'Scroll Bar'], [0x80000005, 'Window'], [0x80000008, 'Window Text'],
  [0x8000000F, 'Button Face'], [0x80000012, 'Button Text'], [0x8000000A, 'Active Border'],
  [0x80000002, 'Active Caption'], [0x80000009, 'Caption Text'], [0x8000000D, 'Highlight'],
  [0x8000000E, 'Highlight Text']
];

let model = null;
let selection = [];
let tool = 'pointer';
let latched = false;
let propMode = 'cat';
let selectedProp = '(Name)';
let clipboard = null;
let drag = null;
let before = null;

const app = document.getElementById('app');
app.innerHTML = [
  '<div class="shell" tabindex="0" id="shell">',
  '<aside class="toolbox"><div class="pane-title">Toolbox</div><div class="tools" id="tools"></div></aside>',
  '<main class="surface" id="surface"><div class="form-window" id="window">',
  '<div class="caption" id="caption"><span class="title" id="captionText"></span><span class="closebox">×</span></div>',
  '<div class="client" id="client"><div class="handles" id="handles"></div></div></div></main>',
  '<aside class="props"><div class="pane-title">Properties</div>',
  '<select class="object-combo" id="objects"></select>',
  '<div class="mode-bar"><button type="button" id="modeCat" class="on">Categorized</button><button type="button" id="modeAlpha">Alphabetic</button></div>',
  '<div class="grid" id="grid"></div></aside></div>',
  '<div class="menu hidden" id="menu"></div>',
  '<div class="fontpop hidden" id="fontpop"></div>',
  '<div class="overlay hidden" id="overlay"></div>'
].join('');

const shell = document.getElementById('shell');
const client = document.getElementById('client');
const toolsEl = document.getElementById('tools');
const gridEl = document.getElementById('grid');
const menuEl = document.getElementById('menu');
const overlay = document.getElementById('overlay');

function px(points) { return (Number(points) || 0) * PT; }
function fromPx(pixels) { return pixels / PT; }
function num(value, fallback) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}
function flag(value, fallback) {
  if (value === undefined || value === null || value === '') return fallback;
  if (value === false || value === 0) return false;
  return true;
}
function clone(value) { return JSON.parse(JSON.stringify(value)); }
function newKey() { return 'k' + Math.random().toString(16).slice(2, 10); }

function propsOf(node) { return node.properties || (node.properties = {}); }
function gridX() { return num(propsOf(model).GridX, 6) || 6; }
function gridY() { return num(propsOf(model).GridY, 6) || 6; }
function snapOn() { return flag(propsOf(model).SnapToGrid, true); }
function showGrid() {
  const p = propsOf(model);
  if (Object.prototype.hasOwnProperty.call(p, 'ShowGridDots')) return flag(p.ShowGridDots, true);
  return flag(p.ShowGrid, true);
}
function clientSize() {
  const p = propsOf(model);
  return {
    width: num(p.InsideWidth, num(p.Width, 240)),
    height: num(p.InsideHeight, num(p.Height, 180))
  };
}

function eachControl(node, visit, parent) {
  for (const child of (node && node.children) || []) {
    visit(child, parent || null);
    eachControl(child, visit, child);
  }
}

function findControl(key) {
  let found = null;
  eachControl(model, (child) => { if (child.key === key) found = child; });
  return found;
}

function parentOf(key) {
  let found = null;
  function walk(node) {
    const kids = node.children || [];
    for (let i = 0; i < kids.length; i++) {
      if (kids[i].key === key) found = { parent: node, index: i };
      else walk(kids[i]);
    }
  }
  if (model) walk(model);
  return found;
}

function usedNames() {
  const names = new Set();
  if (model && model.name) names.add(model.name.toLowerCase());
  eachControl(model, (child) => { if (child.name) names.add(child.name.toLowerCase()); });
  return names;
}

function primary() {
  return selection.length ? findControl(selection[selection.length - 1]) : null;
}

function showOverlay(html) {
  overlay.classList.remove('hidden');
  overlay.innerHTML = '<div class="card">' + html + '</div>';
}
function hideOverlay() { overlay.classList.add('hidden'); }

function status(text) { showOverlay(escapeHtml(text)); }

function escapeHtml(text) {
  return String(text).replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
}

function paintTools() {
  toolsEl.innerHTML = TOOLS.map(([id, title, icon]) => {
    const cls = id === tool ? (latched && id !== 'pointer' ? 'tool latched' : 'tool armed') : 'tool';
    return '<button type="button" class="' + cls + '" data-tool="' + id + '" title="' + title + '">' + icon() + '</button>';
  }).join('');
}

function paintForm() {
  if (!model) return;
  const size = clientSize();
  const p = propsOf(model);
  document.getElementById('captionText').textContent = p.Caption || model.name || 'UserForm';
  client.style.width = px(size.width) + 'px';
  client.style.height = px(size.height) + 'px';
  client.style.backgroundColor = cssColor(p.BackColor, '#f0f0f0');
  client.classList.toggle('grid', showGrid());
  client.style.backgroundSize = px(gridX()) + 'px ' + px(gridY()) + 'px';
  [...client.querySelectorAll(':scope > .ctrl')].forEach((el) => el.remove());
  for (const child of model.children || []) client.appendChild(renderControl(child));
  paintObjects();
}

function renderControl(control) {
  const el = document.createElement('div');
  el.className = 'ctrl ctl-' + String(control.type || 'unknown').toLowerCase() + (control.opaque ? ' opaque' : '');
  if (selection.includes(control.key)) el.classList.add('selected');
  el.dataset.key = control.key;
  const p = propsOf(control);
  el.style.left = px(p.Left) + 'px';
  el.style.top = px(p.Top) + 'px';
  el.style.width = px(p.Width || 72) + 'px';
  el.style.height = px(p.Height || 18) + 'px';
  el.style.color = cssColor(p.ForeColor, '#000');
  el.style.font = cssFont(p.Font);
  if (p.Font && p.Font.underline) el.style.textDecoration = 'underline';
  if (control.opaque) {
    el.textContent = control.name || 'ActiveX';
    return el;
  }
  fillControl(el, control);
  return el;
}

function fillControl(el, control) {
  const p = propsOf(control);
  const type = control.type;
  const caption = p.Caption != null ? String(p.Caption) : (control.name || '');
  if (type === 'CommandButton' || type === 'ToggleButton' || type === 'Label') {
    el.textContent = caption;
    if (type === 'ToggleButton' && flag(p.Value, false)) el.classList.add('on');
    if (p.BackColor != null && type !== 'Label') el.style.background = cssColor(p.BackColor, '#f0f0f0');
    if (type === 'Label' && flag(p.BackStyle, 1) === false) el.style.background = 'transparent';
    return;
  }
  if (type === 'TextBox') {
    el.textContent = p.Text != null ? String(p.Text) : '';
    return;
  }
  if (type === 'ComboBox') {
    el.innerHTML = '<span>' + escapeHtml(p.Text != null ? String(p.Text) : '') + '</span><span class="drop">▾</span>';
    return;
  }
  if (type === 'ListBox') {
    el.textContent = p.Text != null ? String(p.Text) : '';
    return;
  }
  if (type === 'CheckBox' || type === 'OptionButton') {
    const on = flag(p.Value, false);
    const mark = type === 'CheckBox'
      ? '<span class="tick">' + (on ? '✓' : '') + '</span>'
      : '<span class="radio' + (on ? ' on' : '') + '"></span>';
    el.innerHTML = mark + '<span>' + escapeHtml(caption) + '</span>';
    return;
  }
  if (type === 'Frame') {
    el.innerHTML = '<span class="frame-cap">' + escapeHtml(caption) + '</span>';
    for (const child of control.children || []) el.appendChild(renderControl(child));
    return;
  }
  if (type === 'Image') {
    el.textContent = 'Image';
    return;
  }
  if (type === 'ScrollBar' || type === 'SpinButton') {
    const horizontal = num(p.Width, 0) > num(p.Height, 1);
    if (horizontal) el.classList.add('horizontal');
    el.innerHTML = '<span class="arrowbtn">' + (horizontal ? '◀' : '▲') + '</span><span class="arrowbtn">' + (horizontal ? '▶' : '▼') + '</span>';
    return;
  }
  if (type === 'MultiPage' || type === 'TabStrip') {
    const pages = type === 'MultiPage' ? (control.children || []) : (control.tabs || [{ caption: 'Tab1' }, { caption: 'Tab2' }]);
    const active = Math.max(0, num(p.Value, 0));
    const tabs = document.createElement('div');
    tabs.className = 'tabs';
    pages.forEach((page, index) => {
      const tab = document.createElement('span');
      tab.className = 'tab' + (index === active ? ' on' : '');
      tab.dataset.tab = String(index);
      tab.dataset.key = control.key;
      const cap = page.properties && page.properties.Caption != null ? page.properties.Caption : (page.caption || page.name || ('Page' + (index + 1)));
      tab.textContent = cap;
      tabs.appendChild(tab);
    });
    el.appendChild(tabs);
    if (type === 'MultiPage') {
      const body = document.createElement('div');
      body.className = 'page-body';
      const page = pages[active] || pages[0];
      if (page) for (const child of page.children || []) body.appendChild(renderControl(child));
      el.appendChild(body);
    }
    return;
  }
  el.textContent = control.name || type || '';
}

function paintObjects() {
  const select = document.getElementById('objects');
  const options = ['<option value="">' + escapeHtml(model.name || 'UserForm') + ' UserForm</option>'];
  eachControl(model, (child) => {
    if (child.type === 'Page') return;
    options.push('<option value="' + escapeHtml(child.key) + '">' + escapeHtml(child.name || child.type) + ' ' + escapeHtml(child.type || '') + '</option>');
  });
  const current = selection.length === 1 ? selection[0] : '';
  select.innerHTML = options.join('');
  select.value = current;
}

function paintHandles() {
  const layer = document.getElementById('handles');
  layer.innerHTML = '';
  if (selection.length !== 1) return;
  const el = client.querySelector('[data-key="' + CSS.escape(selection[0]) + '"]');
  if (!el) return;
  const host = client;
  const er = el.getBoundingClientRect();
  const hr = host.getBoundingClientRect();
  const x = er.left - hr.left;
  const y = er.top - hr.top;
  const w = er.width;
  const h = er.height;
  const spots = {
    nw: [x, y], n: [x + w / 2, y], ne: [x + w, y],
    e: [x + w, y + h / 2], se: [x + w, y + h], s: [x + w / 2, y + h],
    sw: [x, y + h], w: [x, y + h / 2]
  };
  for (const [name, point] of Object.entries(spots)) {
    const handle = document.createElement('div');
    handle.className = 'handle';
    handle.dataset.handle = name;
    handle.style.left = point[0] + 'px';
    handle.style.top = point[1] + 'px';
    handle.style.cursor = name + '-resize';
    layer.appendChild(handle);
  }
}

function targetNode() {
  if (selection.length === 1) return findControl(selection[0]) || model;
  return model;
}

function paintProps() {
  if (!model) return;
  const node = selection.length > 1 ? null : targetNode();
  if (!node) {
    gridEl.innerHTML = '<div class="cathead">' + selection.length + ' controls selected</div>';
    return;
  }
  const rows = propertyRows(node);
  if (propMode === 'alpha') {
    rows.sort((a, b) => a.name === '(Name)' ? -1 : b.name === '(Name)' ? 1 : a.name.localeCompare(b.name));
    gridEl.innerHTML = rows.map(renderRow).join('');
    return;
  }
  const order = ['Appearance', 'Behavior', 'Data', 'Font', 'Misc', 'Position', 'Scrolling'];
  let html = '';
  for (const cat of order) {
    const group = rows.filter((row) => row.cat === cat);
    if (!group.length) continue;
    html += '<div class="cathead">' + cat + '</div>' + group.map(renderRow).join('');
  }
  const rest = rows.filter((row) => !order.includes(row.cat));
  if (rest.length) html += '<div class="cathead">Misc</div>' + rest.map(renderRow).join('');
  gridEl.innerHTML = html;
}

function propertyRows(node) {
  const form = node === model;
  const p = propsOf(node);
  const rows = [{ name: '(Name)', cat: 'Misc', kind: 'name', value: form ? model.name : node.name, readonly: form }];
  const seen = new Set(['Name', '(Name)']);
  Object.keys(p).forEach((name) => {
    if (seen.has(name)) return;
    seen.add(name);
    const value = p[name];
    if (value && typeof value === 'object' && value.picture) {
      rows.push({ name, cat: CATEGORY[name] || 'Appearance', kind: 'picture', value: '(Binary)', readonly: true });
      return;
    }
    if (name === 'Font' || (value && typeof value === 'object')) {
      rows.push({ name, cat: 'Font', kind: 'font', value });
      return;
    }
    let kind = 'text';
    if (ENUMS[name]) kind = 'enum';
    else if (/Color$/.test(name)) kind = 'color';
    else if (typeof value === 'boolean' || name === 'SnapToGrid' || name === 'ShowGridDots' || name === 'ShowModal' || name === 'Enabled' || name === 'Visible') kind = 'bool';
    else if (typeof value === 'number') kind = 'number';
    rows.push({ name, cat: CATEGORY[name] || 'Misc', kind, value });
  });
  return rows;
}

function renderRow(row) {
  const sel = row.name === selectedProp ? ' sel' : '';
  return '<div class="prow' + sel + '" data-prop="' + escapeHtml(row.name) + '"><div class="pname">' + escapeHtml(row.name) + '</div><div class="pval">' + renderValue(row) + '</div></div>';
}

function renderValue(row) {
  if (row.readonly || row.kind === 'picture') return escapeHtml(row.value == null ? '' : String(row.value));
  if (row.kind === 'bool') {
    const on = flag(row.value, false);
    return '<select data-edit="' + escapeHtml(row.name) + '"><option value="true"' + (on ? ' selected' : '') + '>True</option><option value="false"' + (!on ? ' selected' : '') + '>False</option></select>';
  }
  if (row.kind === 'enum') {
    const options = ENUMS[row.name].map(([value, label]) => {
      const selected = Number(row.value) === Number(value) ? ' selected' : '';
      return '<option value="' + value + '"' + selected + '>' + escapeHtml(label) + '</option>';
    }).join('');
    return '<select data-edit="' + escapeHtml(row.name) + '">' + options + '</select>';
  }
  if (row.kind === 'color') {
    const shown = formatOle(row.value);
    const sys = SYSTEM_COLORS.map(([value, label]) => '<option value="' + value + '">' + escapeHtml(label) + '</option>').join('');
    return '<input data-edit="' + escapeHtml(row.name) + '" value="' + escapeHtml(shown) + '"><select data-color="' + escapeHtml(row.name) + '"><option value="">System…</option>' + sys + '</select>';
  }
  if (row.kind === 'font') {
    const font = row.value || {};
    const label = (font.name || 'Tahoma') + ', ' + (font.size || 8) + 'pt';
    return '<button type="button" data-font="1">' + escapeHtml(label) + ' …</button>';
  }
  const shown = row.value == null ? '' : String(row.value);
  return '<input data-edit="' + escapeHtml(row.name) + '" value="' + escapeHtml(shown) + '">';
}

function cssColor(value, fallback) {
  if (value === undefined || value === null || value === '') return fallback;
  const n = Number(value) >>> 0;
  if (!Number.isFinite(Number(value))) return fallback;
  const named = {
    0x80000005: '#ffffff', 0x80000008: '#000000', 0x8000000f: '#f0f0f0',
    0x80000012: '#000000', 0x80000002: '#0a246a', 0x80000009: '#ffffff',
    0x8000000a: '#d4d0c8', 0x8000000d: '#0078d7', 0x8000000e: '#ffffff'
  };
  if ((n & 0x80000000) !== 0) return named[n] || fallback;
  return 'rgb(' + (n & 255) + ',' + ((n >> 8) & 255) + ',' + ((n >> 16) & 255) + ')';
}

function cssFont(font) {
  if (!font || typeof font !== 'object') return '8pt Tahoma, sans-serif';
  return (font.italic ? 'italic ' : '') + (font.bold ? 'bold ' : '') + (font.size || 8) + 'pt ' + (font.name || 'Tahoma') + ', sans-serif';
}

function formatOle(value) {
  const n = Number(value) >>> 0;
  return '&H' + n.toString(16).toUpperCase().padStart(8, '0') + '&';
}

function parseOle(text) {
  const raw = String(text).trim();
  const hex = raw.match(/^&H([0-9A-Fa-f]+)&$/i);
  if (hex) return parseInt(hex[1], 16);
  if (/^-?\d+$/.test(raw)) return Number(raw);
  return null;
}

function commit(next) {
  if (!model) return;
  const previous = before || clone(model);
  model = next;
  before = null;
  if (JSON.stringify(previous) === JSON.stringify(model)) return;
  vscode.postMessage({ type: 'edit', before: previous, after: clone(model) });
  render();
}

function mutate(fn) {
  const next = clone(model);
  fn(next);
  commit(next);
}

function beginGesture() {
  if (!before) before = clone(model);
}

function endGesture() {
  if (!before) return;
  const previous = before;
  before = null;
  if (JSON.stringify(previous) === JSON.stringify(model)) return;
  vscode.postMessage({ type: 'edit', before: previous, after: clone(model) });
}

function selectOnly(key) { selection = key ? [key] : []; }
function render() { paintTools(); paintForm(); paintProps(); paintHandles(); }

function makeControl(type, left, top, width, height) {
  const names = usedNames();
  const name = edit.nextName(type, names);
  const control = {
    key: newKey(),
    name,
    progId: 'Forms.' + type + '.1',
    type,
    opaque: false,
    properties: { Left: left, Top: top, Width: width, Height: height },
    children: []
  };
  if (['Label', 'CommandButton', 'CheckBox', 'OptionButton', 'ToggleButton', 'Frame'].includes(type)) {
    control.properties.Caption = name;
  }
  if (type === 'MultiPage') {
    control.properties.Value = 0;
    control.children = [1, 2].map((n) => ({
      key: newKey(),
      name: 'Page' + n,
      progId: 'Forms.Page.1',
      type: 'Page',
      properties: { Caption: 'Page' + n },
      children: []
    }));
  }
  if (type === 'TabStrip') {
    control.properties.Value = 0;
    control.tabs = [{ name: 'Tab1', caption: 'Tab1' }, { name: 'Tab2', caption: 'Tab2' }];
  }
  return control;
}

function dropTarget(event) {
  const el = event.target.closest && event.target.closest('.ctrl');
  if (!el) return { parent: model, box: client };
  const control = findControl(el.dataset.key);
  if (control && (control.type === 'Frame' || control.type === 'Page')) return { parent: control, box: el };
  const pageBody = event.target.closest && event.target.closest('.page-body');
  if (pageBody) {
    const host = pageBody.parentElement;
    const multi = findControl(host.dataset.key);
    const pages = (multi && multi.children) || [];
    const page = pages[num(propsOf(multi).Value, 0)] || pages[0];
    if (page) return { parent: page, box: pageBody };
  }
  return { parent: model, box: client };
}

function pointIn(event, box) {
  const rect = box.getBoundingClientRect();
  return {
    x: fromPx(event.clientX - rect.left - box.clientLeft),
    y: fromPx(event.clientY - rect.top - box.clientTop)
  };
}

toolsEl.addEventListener('click', (event) => {
  const button = event.target.closest('[data-tool]');
  if (!button) return;
  const id = button.dataset.tool;
  if (id === 'pointer') { tool = 'pointer'; latched = false; }
  else if (tool === id && latched) { tool = 'pointer'; latched = false; }
  else { tool = id; latched = false; }
  paintTools();
});

toolsEl.addEventListener('dblclick', (event) => {
  const button = event.target.closest('[data-tool]');
  if (!button || button.dataset.tool === 'pointer') return;
  tool = button.dataset.tool;
  latched = true;
  paintTools();
});

client.addEventListener('mousedown', onSurfaceDown);
document.getElementById('caption').addEventListener('mousedown', (event) => {
  if (event.detail === 2) return;
  selection = [];
  render();
});
document.getElementById('caption').addEventListener('dblclick', () => {
  vscode.postMessage({ type: 'viewCode', name: model.name, event: 'Click' });
});

function onSurfaceDown(event) {
  if (event.button !== 0) return;
  menuEl.classList.add('hidden');
  const tab = event.target.closest && event.target.closest('.tab');
  if (tab) {
    const control = findControl(tab.dataset.key);
    if (control) {
      beginGesture();
      propsOf(control).Value = Number(tab.dataset.tab);
      endGesture();
      selectOnly(control.key);
      render();
    }
    return;
  }
  const hit = event.target.closest && event.target.closest('.ctrl');
  if (tool !== 'pointer') {
    const where = dropTarget(event);
    const origin = pointIn(event, where.box);
    drag = { kind: 'draw', origin, where, x: origin.x, y: origin.y };
    beginGesture();
    event.preventDefault();
    return;
  }
  if (!hit) {
    drag = { kind: 'marquee', x: event.clientX, y: event.clientY, sx: event.clientX, sy: event.clientY };
    if (!event.shiftKey) selection = [];
    beginGesture();
    render();
    return;
  }
  const key = hit.dataset.key;
  const control = findControl(key);
  if (!control || control.opaque) {
    selectOnly(key);
    render();
    return;
  }
  if (event.shiftKey) {
    selection = selection.includes(key) ? selection.filter((item) => item !== key) : selection.concat(key);
  } else if (!selection.includes(key)) {
    selectOnly(key);
  } else {
    selection = selection.filter((item) => item !== key).concat(key);
  }
  const start = {};
  selection.forEach((item) => {
    const node = findControl(item);
    if (!node) return;
    start[item] = { Left: num(propsOf(node).Left, 0), Top: num(propsOf(node).Top, 0) };
  });
  drag = { kind: 'move', x: event.clientX, y: event.clientY, start, key };
  beginGesture();
  render();
  event.preventDefault();
}

document.getElementById('handles').addEventListener('mousedown', (event) => {
  const handle = event.target.closest('[data-handle]');
  if (!handle || selection.length !== 1) return;
  const control = findControl(selection[0]);
  if (!control) return;
  const p = propsOf(control);
  drag = {
    kind: 'resize',
    handle: handle.dataset.handle,
    x: event.clientX,
    y: event.clientY,
    Left: num(p.Left, 0),
    Top: num(p.Top, 0),
    Width: num(p.Width, 12),
    Height: num(p.Height, 12)
  };
  beginGesture();
  event.preventDefault();
  event.stopPropagation();
});

window.addEventListener('mousemove', (event) => {
  if (!drag) return;
  if (drag.kind === 'move') {
    const dx = fromPx(event.clientX - drag.x);
    const dy = fromPx(event.clientY - drag.y);
    selection.forEach((key) => {
      const node = findControl(key);
      const origin = drag.start[key];
      if (!node || !origin || node.opaque) return;
      propsOf(node).Left = edit.snap(origin.Left + dx, gridX(), snapOn());
      propsOf(node).Top = edit.snap(origin.Top + dy, gridY(), snapOn());
    });
    paintForm();
    paintHandles();
  } else if (drag.kind === 'resize') {
    const node = findControl(selection[0]);
    if (!node) return;
    const dx = fromPx(event.clientX - drag.x);
    const dy = fromPx(event.clientY - drag.y);
    let left = drag.Left;
    let top = drag.Top;
    let width = drag.Width;
    let height = drag.Height;
    const h = drag.handle;
    if (h.includes('e')) width = drag.Width + dx;
    if (h.includes('s')) height = drag.Height + dy;
    if (h.includes('w')) { width = drag.Width - dx; left = drag.Left + dx; }
    if (h.includes('n')) { height = drag.Height - dy; top = drag.Top + dy; }
    width = Math.max(4, edit.snap(width, gridX(), snapOn()));
    height = Math.max(4, edit.snap(height, gridY(), snapOn()));
    propsOf(node).Left = edit.snap(left, gridX(), snapOn());
    propsOf(node).Top = edit.snap(top, gridY(), snapOn());
    propsOf(node).Width = width;
    propsOf(node).Height = height;
    paintForm();
    paintHandles();
  } else if (drag.kind === 'draw') {
    const box = drag.where.box;
    const point = pointIn(event, box);
    showMarquee(drag.origin, point, box);
  } else if (drag.kind === 'marquee') {
    showScreenMarquee(drag.sx, drag.sy, event.clientX, event.clientY);
  }
});

window.addEventListener('mouseup', (event) => {
  if (!drag) return;
  const action = drag;
  drag = null;
  clearRubber();
  if (action.kind === 'draw') {
    const box = action.where.box;
    const point = pointIn(event, box);
    let left = Math.min(action.origin.x, point.x);
    let top = Math.min(action.origin.y, point.y);
    let width = Math.abs(point.x - action.origin.x);
    let height = Math.abs(point.y - action.origin.y);
    const size = edit.defaultSize(tool);
    if (width < 2 && height < 2) {
      left = action.origin.x;
      top = action.origin.y;
      width = size.width;
      height = size.height;
    }
    left = edit.snap(left, gridX(), snapOn());
    top = edit.snap(top, gridY(), snapOn());
    width = Math.max(4, edit.snap(width, gridX(), snapOn()));
    height = Math.max(4, edit.snap(height, gridY(), snapOn()));
    const created = makeControl(tool, left, top, width, height);
    action.where.parent.children = action.where.parent.children || [];
    action.where.parent.children.unshift(created);
    selectOnly(created.key);
    if (!latched) tool = 'pointer';
    endGesture();
    render();
    return;
  }
  if (action.kind === 'marquee') {
    const rect = marqueeRect(action.sx, action.sy, event.clientX, event.clientY);
    const hits = [];
    client.querySelectorAll('.ctrl').forEach((el) => {
      const r = el.getBoundingClientRect();
      if (r.left < rect.right && r.right > rect.left && r.top < rect.bottom && r.bottom > rect.top) hits.push(el.dataset.key);
    });
    selection = event.shiftKey ? Array.from(new Set(selection.concat(hits))) : hits;
    endGesture();
    render();
    return;
  }
  endGesture();
  paintProps();
});

function showMarquee(a, b, box) {
  let band = document.getElementById('drawband');
  if (!band) {
    band = document.createElement('div');
    band.id = 'drawband';
    band.className = 'marquee';
    box.appendChild(band);
  }
  const left = Math.min(a.x, b.x);
  const top = Math.min(a.y, b.y);
  band.style.left = px(left) + 'px';
  band.style.top = px(top) + 'px';
  band.style.width = px(Math.abs(b.x - a.x)) + 'px';
  band.style.height = px(Math.abs(b.y - a.y)) + 'px';
}

function showScreenMarquee(x1, y1, x2, y2) {
  let band = document.getElementById('drawband');
  if (!band) {
    band = document.createElement('div');
    band.id = 'drawband';
    band.className = 'marquee';
    document.getElementById('surface').appendChild(band);
  }
  const host = document.getElementById('surface');
  const hr = host.getBoundingClientRect();
  const left = Math.min(x1, x2) - hr.left + host.scrollLeft;
  const top = Math.min(y1, y2) - hr.top + host.scrollTop;
  band.style.left = left + 'px';
  band.style.top = top + 'px';
  band.style.width = Math.abs(x2 - x1) + 'px';
  band.style.height = Math.abs(y2 - y1) + 'px';
}

function marqueeRect(x1, y1, x2, y2) {
  return { left: Math.min(x1, x2), right: Math.max(x1, x2), top: Math.min(y1, y2), bottom: Math.max(y1, y2) };
}

function clearRubber() {
  const band = document.getElementById('drawband');
  if (band) band.remove();
}

shell.addEventListener('keydown', (event) => {
  if (event.target.matches('input, select, textarea')) return;
  const key = event.key;
  if ((event.ctrlKey || event.metaKey) && key.toLowerCase() === 'z') {
    event.preventDefault();
    vscode.postMessage({ type: event.shiftKey ? 'redo' : 'undo' });
    return;
  }
  if ((event.ctrlKey || event.metaKey) && key.toLowerCase() === 'y') {
    event.preventDefault();
    vscode.postMessage({ type: 'redo' });
    return;
  }
  if ((event.ctrlKey || event.metaKey) && key.toLowerCase() === 'c') { event.preventDefault(); copySelection(); return; }
  if ((event.ctrlKey || event.metaKey) && key.toLowerCase() === 'x') { event.preventDefault(); cutSelection(); return; }
  if ((event.ctrlKey || event.metaKey) && key.toLowerCase() === 'v') { event.preventDefault(); pasteSelection(); return; }
  if ((event.ctrlKey || event.metaKey) && key.toLowerCase() === 'd') { event.preventDefault(); copySelection(); pasteSelection(); return; }
  if ((event.ctrlKey || event.metaKey) && key.toLowerCase() === 'a') {
    event.preventDefault();
    selection = [];
    eachControl(model, (child) => { if (child.type !== 'Page') selection.push(child.key); });
    render();
    return;
  }
  if (key === 'Escape') {
    tool = 'pointer';
    latched = false;
    selection = [];
    menuEl.classList.add('hidden');
    render();
    return;
  }
  if (key === 'Delete' || key === 'Backspace') { event.preventDefault(); deleteSelection(); return; }
  if (key === 'F2') { event.preventDefault(); beginCaptionEdit(); return; }
  const arrows = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] };
  if (!arrows[key] || !selection.length) return;
  event.preventDefault();
  const stepX = snapOn() ? gridX() : 1;
  const stepY = snapOn() ? gridY() : 1;
  const [dx, dy] = arrows[key];
  mutate((next) => {
    selection.forEach((id) => {
      let node = null;
      eachControl(next, (child) => { if (child.key === id) node = child; });
      if (!node || node.opaque) return;
      const p = propsOf(node);
      if (event.shiftKey) {
        p.Width = Math.max(4, num(p.Width, 12) + dx * stepX);
        p.Height = Math.max(4, num(p.Height, 12) + dy * stepY);
      } else {
        p.Left = num(p.Left, 0) + dx * stepX;
        p.Top = num(p.Top, 0) + dy * stepY;
      }
    });
  });
});

client.addEventListener('dblclick', (event) => {
  const hit = event.target.closest('.ctrl');
  if (!hit) {
    vscode.postMessage({ type: 'viewCode', name: model.name, event: 'Click' });
    return;
  }
  const control = findControl(hit.dataset.key);
  if (!control) return;
  vscode.postMessage({ type: 'viewCode', name: control.name, event: edit.defaultEvent(control.progId) });
});

client.addEventListener('contextmenu', openMenu);
document.getElementById('caption').addEventListener('contextmenu', openMenu);

function openMenu(event) {
  event.preventDefault();
  const hit = event.target.closest && event.target.closest('.ctrl');
  if (hit && !selection.includes(hit.dataset.key)) {
    selectOnly(hit.dataset.key);
    render();
  }
  const items = [
    ['View Code', () => viewPrimary()],
    ['-', null],
    ['Cut', cutSelection],
    ['Copy', copySelection],
    ['Paste', pasteSelection],
    ['Delete', deleteSelection],
    ['-', null],
    ['Bring to Front', () => order(true)],
    ['Send to Back', () => order(false)],
    ['-', null],
    ['Align Left', () => align('left')],
    ['Align Center', () => align('center')],
    ['Align Right', () => align('right')],
    ['Align Top', () => align('top')],
    ['Align Middle', () => align('middle')],
    ['Align Bottom', () => align('bottom')],
    ['Make Same Width', () => sameSize('width')],
    ['Make Same Height', () => sameSize('height')],
    ['Center in Form Horizontally', () => centerInForm('x')],
    ['Center in Form Vertically', () => centerInForm('y')]
  ];
  menuEl.innerHTML = items.map(([label]) => label === '-' ? '<hr>' : '<button type="button">' + label + '</button>').join('');
  menuEl.classList.remove('hidden');
  menuEl.style.left = Math.min(event.clientX, window.innerWidth - 200) + 'px';
  menuEl.style.top = Math.min(event.clientY, window.innerHeight - 280) + 'px';
  [...menuEl.querySelectorAll('button')].forEach((button, index) => {
    const real = items.filter((item) => item[0] !== '-');
    const match = real.find((item) => item[0] === button.textContent);
    button.addEventListener('click', () => { menuEl.classList.add('hidden'); if (match) match[1](); });
    void index;
  });
}

function viewPrimary() {
  const control = primary();
  if (!control) vscode.postMessage({ type: 'viewCode', name: model.name, event: 'Click' });
  else vscode.postMessage({ type: 'viewCode', name: control.name, event: edit.defaultEvent(control.progId) });
}

function copySelection() {
  clipboard = selection.map(findControl).filter((node) => node && !node.opaque).map((node) => clone(node));
}

function cutSelection() { copySelection(); deleteSelection(); }

function pasteSelection() {
  if (!clipboard || !clipboard.length) return;
  const names = usedNames();
  const step = gridX() || 6;
  const copies = clipboard.map((node) => cloneTree(node, names, step, step));
  mutate((next) => {
    const host = parentNode(next, selection[0]) || next;
    host.children = host.children || [];
    copies.forEach((copy) => host.children.unshift(copy));
  });
  selection = copies.map((copy) => copy.key);
  render();
}

function cloneTree(node, names, dx, dy) {
  const copy = clone(node);
  const type = copy.type === 'Page' ? 'Page' : (copy.type || 'Control');
  copy.key = newKey();
  copy.name = edit.nextName(type, names);
  names.add(copy.name.toLowerCase());
  const p = propsOf(copy);
  if ('Left' in p) p.Left = num(p.Left, 0) + dx;
  if ('Top' in p) p.Top = num(p.Top, 0) + dy;
  if (p.Caption === node.name) p.Caption = copy.name;
  copy.children = (copy.children || []).map((child) => cloneTree(child, names, 0, 0));
  return copy;
}

function parentNode(root, key) {
  if (!key) return root;
  let found = null;
  function walk(node) {
    for (const child of node.children || []) {
      if (child.key === key) found = node;
      else walk(child);
    }
  }
  walk(root);
  return found;
}

function deleteSelection() {
  const doomed = new Set(selection.filter((key) => {
    const node = findControl(key);
    return node && !node.opaque;
  }));
  if (!doomed.size) return;
  mutate((next) => {
    function walk(node) {
      node.children = (node.children || []).filter((child) => !doomed.has(child.key));
      node.children.forEach(walk);
    }
    walk(next);
  });
  selection = [];
  render();
}

function order(front) {
  mutate((next) => {
    selection.forEach((key) => {
      const holder = parentNode(next, key);
      if (!holder) return;
      const index = holder.children.findIndex((child) => child.key === key);
      if (index < 0) return;
      const [item] = holder.children.splice(index, 1);
      if (front) holder.children.unshift(item);
      else holder.children.push(item);
    });
  });
}

function selectedNodes() {
  return selection.map(findControl).filter((node) => node && !node.opaque);
}

function align(kind) {
  const nodes = selectedNodes();
  if (nodes.length < 2) return;
  const lead = nodes[nodes.length - 1];
  const lp = propsOf(lead);
  mutate((next) => {
    nodes.forEach((node) => {
      let live = null;
      eachControl(next, (child) => { if (child.key === node.key) live = child; });
      if (!live) return;
      const p = propsOf(live);
      if (kind === 'left') p.Left = num(lp.Left, 0);
      if (kind === 'right') p.Left = num(lp.Left, 0) + num(lp.Width, 0) - num(p.Width, 0);
      if (kind === 'center') p.Left = num(lp.Left, 0) + num(lp.Width, 0) / 2 - num(p.Width, 0) / 2;
      if (kind === 'top') p.Top = num(lp.Top, 0);
      if (kind === 'bottom') p.Top = num(lp.Top, 0) + num(lp.Height, 0) - num(p.Height, 0);
      if (kind === 'middle') p.Top = num(lp.Top, 0) + num(lp.Height, 0) / 2 - num(p.Height, 0) / 2;
    });
  });
}

function sameSize(axis) {
  const nodes = selectedNodes();
  if (nodes.length < 2) return;
  const lead = propsOf(nodes[nodes.length - 1]);
  mutate((next) => {
    nodes.forEach((node) => {
      let live = null;
      eachControl(next, (child) => { if (child.key === node.key) live = child; });
      if (!live) return;
      if (axis === 'width') propsOf(live).Width = num(lead.Width, 12);
      else propsOf(live).Height = num(lead.Height, 12);
    });
  });
}

function centerInForm(axis) {
  const nodes = selectedNodes();
  if (!nodes.length) return;
  const size = clientSize();
  mutate((next) => {
    nodes.forEach((node) => {
      let live = null;
      eachControl(next, (child) => { if (child.key === node.key) live = child; });
      if (!live) return;
      const p = propsOf(live);
      if (axis === 'x') p.Left = (size.width - num(p.Width, 0)) / 2;
      else p.Top = (size.height - num(p.Height, 0)) / 2;
    });
  });
}

function beginCaptionEdit() {
  const control = primary();
  const node = control || model;
  const p = propsOf(node);
  if (control && !('Caption' in p) && control.type !== 'UserForm') {
    if (!['Label', 'CommandButton', 'CheckBox', 'OptionButton', 'ToggleButton', 'Frame'].includes(control.type)) return;
  }
  const host = control ? client.querySelector('[data-key="' + CSS.escape(control.key) + '"]') : document.getElementById('captionText');
  if (!host) return;
  const input = document.createElement('input');
  input.className = 'inplace';
  input.value = p.Caption != null ? String(p.Caption) : (node.name || '');
  const rect = host.getBoundingClientRect();
  input.style.left = rect.left + 'px';
  input.style.top = rect.top + 'px';
  input.style.width = Math.max(40, rect.width) + 'px';
  document.body.appendChild(input);
  input.focus();
  input.select();
  const finish = (save) => {
    const value = input.value;
    input.remove();
    if (!save) return;
    mutate((next) => {
      const live = control ? (function () { let f = null; eachControl(next, (child) => { if (child.key === control.key) f = child; }); return f; })() : next;
      if (live) propsOf(live).Caption = value;
    });
  };
  input.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') { event.preventDefault(); finish(true); }
    if (event.key === 'Escape') { event.preventDefault(); finish(false); }
  });
  input.addEventListener('blur', () => finish(true));
}

document.getElementById('objects').addEventListener('change', (event) => {
  selectOnly(event.target.value);
  render();
});
document.getElementById('modeCat').addEventListener('click', () => { propMode = 'cat'; paintProps(); document.getElementById('modeCat').classList.add('on'); document.getElementById('modeAlpha').classList.remove('on'); });
document.getElementById('modeAlpha').addEventListener('click', () => { propMode = 'alpha'; paintProps(); document.getElementById('modeAlpha').classList.add('on'); document.getElementById('modeCat').classList.remove('on'); });

gridEl.addEventListener('mousedown', (event) => {
  const row = event.target.closest('.prow');
  if (!row) return;
  selectedProp = row.dataset.prop;
  paintProps();
});

gridEl.addEventListener('change', (event) => {
  const editEl = event.target.closest('[data-edit], [data-color]');
  if (!editEl) return;
  const name = editEl.dataset.edit || editEl.dataset.color;
  let value = editEl.value;
  if (editEl.dataset.color) {
    if (value === '') return;
    value = Number(value);
  } else if (ENUMS[name] || name === 'Value' && /^-?\d+$/.test(value)) {
    value = value === 'true' ? true : value === 'false' ? false : Number(value);
  } else if (value === 'true' || value === 'false') {
    value = value === 'true';
  } else if (/Color$/.test(name)) {
    const parsed = parseOle(value);
    if (parsed === null) return;
    value = parsed;
  } else if (value !== '' && /^-?\d+(\.\d+)?$/.test(value) && !['PasswordChar', 'Caption', 'Text', 'Tag', 'ControlTipText', 'Accelerator', 'ControlSource', 'RowSource', 'GroupName', '(Name)'].includes(name)) {
    value = Number(value);
  }
  applyProp(name, value);
});

gridEl.addEventListener('click', (event) => {
  const font = event.target.closest('[data-font]');
  if (!font) return;
  openFont(font);
});

function applyProp(name, value) {
  if (name === '(Name)') {
    const text = String(value).trim();
    if (!/^[A-Za-z][A-Za-z0-9_]{0,30}$/.test(text)) return;
    if (usedNames().has(text.toLowerCase()) && primary() && primary().name.toLowerCase() !== text.toLowerCase()) return;
    mutate((next) => {
      let live = null;
      eachControl(next, (child) => { if (selection.includes(child.key) && child.key === selection[selection.length - 1]) live = child; });
      if (live) live.name = text;
    });
    return;
  }
  mutate((next) => {
    const live = selection.length ? (function () { let f = null; eachControl(next, (child) => { if (child.key === selection[selection.length - 1]) f = child; }); return f; })() : next;
    if (!live) return;
    propsOf(live)[name] = value;
  });
}

function openFont(anchor) {
  const node = targetNode();
  const font = Object.assign({ name: 'Tahoma', size: 8, bold: false, italic: false, underline: false }, propsOf(node).Font || {});
  const pop = document.getElementById('fontpop');
  pop.classList.remove('hidden');
  const rect = anchor.getBoundingClientRect();
  pop.style.left = rect.left + 'px';
  pop.style.top = (rect.bottom + 2) + 'px';
  pop.innerHTML = [
    '<label>Name <input id="fn" value="' + escapeHtml(font.name) + '"></label>',
    '<label>Size <input id="fs" value="' + escapeHtml(font.size) + '"></label>',
    '<label><input type="checkbox" id="fb"' + (font.bold ? ' checked' : '') + '> Bold</label>',
    '<label><input type="checkbox" id="fi"' + (font.italic ? ' checked' : '') + '> Italic</label>',
    '<label><input type="checkbox" id="fu"' + (font.underline ? ' checked' : '') + '> Underline</label>',
    '<button type="button" id="fok">OK</button>'
  ].join('');
  document.getElementById('fok').addEventListener('click', () => {
    applyProp('Font', {
      name: document.getElementById('fn').value || 'Tahoma',
      size: Number(document.getElementById('fs').value) || 8,
      bold: document.getElementById('fb').checked,
      italic: document.getElementById('fi').checked,
      underline: document.getElementById('fu').checked
    });
    pop.classList.add('hidden');
  });
}

window.addEventListener('message', (event) => {
  const msg = event.data || {};
  if (msg.type === 'status') { status(msg.text || ''); return; }
  if (msg.type === 'error') {
    showOverlay('<b>' + escapeHtml(msg.message || 'Could not open the form') + '</b><br><br>' + escapeHtml(msg.remedy || ''));
    return;
  }
  if (msg.type === 'model' && msg.model) {
    model = msg.model;
    if (!Array.isArray(selection)) selection = [];
    selection = selection.filter((key) => findControl(key));
    hideOverlay();
    render();
    shell.focus();
  }
});

document.getElementById('surface').addEventListener('scroll', paintHandles);
window.addEventListener('resize', paintHandles);

function pointerIcon() { return '<svg viewBox="0 0 16 16"><path d="M3 1 L3 13 L6 10 L8 14 L10 13 L8 9 L12 9 Z" fill="#000"/></svg>'; }
function labelIcon() { return '<svg viewBox="0 0 16 16"><text x="3" y="13" font-size="13" font-family="Tahoma" fill="#000">A</text></svg>'; }
function textIcon() { return '<svg viewBox="0 0 16 16"><text x="1" y="12" font-size="10" font-family="Tahoma" fill="#000">ab</text><rect x="12" y="3" width="1" height="10" fill="#000"/></svg>'; }
function comboIcon() { return '<svg viewBox="0 0 16 16"><rect x="1" y="3" width="14" height="10" fill="#fff" stroke="#000"/><rect x="10" y="3" width="5" height="10" fill="#f0f0f0" stroke="#000"/><path d="M11 7 L13 10 L15 7" fill="none" stroke="#000"/></svg>'; }
function listIcon() { return '<svg viewBox="0 0 16 16"><rect x="2" y="2" width="12" height="12" fill="#fff" stroke="#000"/><path d="M4 5 H12 M4 8 H12 M4 11 H10" stroke="#000"/></svg>'; }
function checkIcon() { return '<svg viewBox="0 0 16 16"><rect x="1" y="3" width="9" height="9" fill="#fff" stroke="#000"/><path d="M3 8 L5 10 L8 5" fill="none" stroke="#000"/><path d="M11 6 H15 M11 9 H15" stroke="#000"/></svg>'; }
function optionIcon() { return '<svg viewBox="0 0 16 16"><circle cx="5" cy="8" r="4" fill="#fff" stroke="#000"/><circle cx="5" cy="8" r="1.6" fill="#000"/><path d="M11 6 H15 M11 9 H15" stroke="#000"/></svg>'; }
function toggleIcon() { return '<svg viewBox="0 0 16 16"><rect x="2" y="4" width="12" height="8" fill="#d4d0c8" stroke="#000"/></svg>'; }
function frameIcon() { return '<svg viewBox="0 0 16 16"><path d="M2 5 H5 M9 5 H14 V13 H2 V5" fill="none" stroke="#000"/></svg>'; }
function buttonIcon() { return '<svg viewBox="0 0 16 16"><rect x="2" y="4" width="12" height="8" fill="#f0f0f0" stroke="#808080"/><path d="M2 4 H14 V5 H2 Z" fill="#fff"/></svg>'; }
function tabsIcon() { return '<svg viewBox="0 0 16 16"><path d="M2 6 H6 V4 H10 V6 H14 V13 H2 Z" fill="#f0f0f0" stroke="#000"/></svg>'; }
function pageIcon() { return '<svg viewBox="0 0 16 16"><path d="M2 5 H7 V3 H12 V5 H14 V13 H2 Z" fill="#fff" stroke="#000"/></svg>'; }
function scrollIcon() { return '<svg viewBox="0 0 16 16"><rect x="6" y="1" width="4" height="14" fill="#f0f0f0" stroke="#000"/><path d="M7 3 L8 1.5 L9 3 M7 13 L8 14.5 L9 13" fill="#000"/></svg>'; }
function spinIcon() { return '<svg viewBox="0 0 16 16"><rect x="5" y="2" width="6" height="12" fill="#f0f0f0" stroke="#000"/><path d="M6 6 L8 4 L10 6 M6 10 L8 12 L10 10" fill="none" stroke="#000"/></svg>'; }
function imageIcon() { return '<svg viewBox="0 0 16 16"><rect x="1" y="2" width="14" height="12" fill="#fff" stroke="#000"/><path d="M2 12 L6 7 L9 10 L11 8 L14 12" fill="none" stroke="#000"/><circle cx="5" cy="5" r="1" fill="#000"/></svg>'; }

vscode.postMessage({ type: 'ready' });
