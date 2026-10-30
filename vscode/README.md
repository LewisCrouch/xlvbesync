# xlvbesync

Go to definition, member completion, and hover docs for extracted VBA, plus the Excel and VBA built-ins. UserForms open in a designer that follows the Excel VBE: toolbox, dotted grid, classic MSForms controls, and a property grid.

Syntax highlighting still comes from the editor you already use. This extension contributes `.dcm` as the `vba` language and also serves the `vb` language this repo maps `.bas`, `.cls`, and `.frm` to.

Do not enable VBA Pro, xlcraft, or XLIDE at the same time. Their definition providers will fight this one.

## Install (Cursor)

From the repo root, junction the folder into Cursor's extensions directory, then run **Developer: Reload Window**:

```powershell
$dest = Join-Path $env:USERPROFILE ".cursor\extensions\xlvbesync.xlvbesync-0.1.0"
if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
cmd /c mklink /J "$dest" "$PWD\vscode"
```

The designer drives a private Excel process through `xlvbesync form-host`. Install the Python package (`pip install -e .`) and turn on **Trust access to the VBA project object model** first. `xlvbesync doctor` checks that setting.

Opening a `.frm` shows the designer. **View Code** (editor title, or double-click a control) opens the code-behind in the text editor. A double-click also inserts the default event procedure when it is missing.

`.frx` files stay binary. The designer writes them by asking Excel to export the form. Commit the `.frm` and the `.frx` together.

Indexer checks, with no Excel required:

```powershell
node vscode/test.js
```

Built-in hover data is generated from the installed Excel, VBA, and Office type libraries. To regenerate:

```powershell
powershell -File vscode/scripts/dump-typelibs.ps1
```
