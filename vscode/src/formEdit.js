'use strict';

(function (root, factory) {
  const api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.formEdit = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  const DEFAULT_SIZE = {
    Label: { width: 72, height: 18 },
    TextBox: { width: 72, height: 18 },
    ComboBox: { width: 72, height: 18 },
    ListBox: { width: 72, height: 65.25 },
    CheckBox: { width: 108, height: 18 },
    OptionButton: { width: 108, height: 18 },
    ToggleButton: { width: 72, height: 18 },
    Frame: { width: 144, height: 96 },
    CommandButton: { width: 72, height: 24 },
    TabStrip: { width: 144, height: 108 },
    MultiPage: { width: 144, height: 108 },
    ScrollBar: { width: 18, height: 72 },
    SpinButton: { width: 12.75, height: 25.5 },
    Image: { width: 72, height: 72 }
  };

  const CHANGE_EVENTS = {
    TextBox: true,
    ComboBox: true,
    ListBox: true,
    ScrollBar: true,
    SpinButton: true,
    MultiPage: true,
    TabStrip: true
  };

  function controlType(progId) {
    const match = String(progId || '').match(/Forms\.([A-Za-z]+)/i);
    return match ? match[1] : '';
  }

  function snap(value, grid, enabled) {
    const g = Number(grid);
    const n = Number(value);
    if (!enabled || !g || !Number.isFinite(n)) return n;
    return Math.round(n / g) * g;
  }

  function defaultEvent(progId) {
    const type = controlType(progId);
    return CHANGE_EVENTS[type] ? 'Change' : 'Click';
  }

  function eventStub(name, eventName) {
    return 'Private Sub ' + name + '_' + eventName + '()\r\n    \r\nEnd Sub\r\n';
  }

  function defaultSize(typeName) {
    return DEFAULT_SIZE[typeName] || { width: 72, height: 24 };
  }

  function nextName(typeName, usedNames) {
    const used = usedNames instanceof Set ? usedNames : new Set(usedNames || []);
    const base = typeName || 'Control';
    let n = 1;
    while (used.has((base + n).toLowerCase())) n += 1;
    return base + n;
  }

  function flatten(node, parentKey, acc) {
    const list = acc || [];
    for (const child of (node && node.children) || []) {
      list.push({ control: child, parentKey: parentKey || '' });
      if (child.children && child.children.length) flatten(child, child.key, list);
    }
    return list;
  }

  function walkControls(model) {
    return flatten(model, '', []).map((row) => row.control).filter((control) => {
      return control && control.name && control.type !== 'Page' && !control.opaque;
    }).map((control) => ({
      name: control.name,
      progId: control.progId || '',
      type: control.type || controlType(control.progId)
    }));
  }

  return {
    DEFAULT_SIZE,
    controlType,
    snap,
    defaultEvent,
    eventStub,
    defaultSize,
    nextName,
    flatten,
    walkControls
  };
});
