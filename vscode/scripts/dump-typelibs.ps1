# Regenerates vscode/data/typelibs.json from the installed Excel / VBA / Office type libraries.
# Usage: powershell -File vscode/scripts/dump-typelibs.ps1

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$outFile = Join-Path $repoRoot 'vscode\data\typelibs.json'

$cs = @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using CT = System.Runtime.InteropServices.ComTypes;

public static class TypeLibDump {
  [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
  static extern void LoadTypeLibEx(string szFile, int regKind, out CT.ITypeLib pptlib);

  const int PARAMFLAG_FRETVAL = 0x8;
  const int PARAMFLAG_FOPT = 0x10;
  const int PARAMFLAG_FHASDEFAULT = 0x20;
  const int PARAMFLAG_FLCID = 0x4;
  const int FUNCFLAG_FRESTRICTED = 0x1;
  const int IMPLTYPEFLAG_FDEFAULT = 1;
  const int IMPLTYPEFLAG_FSOURCE = 2;
  const int VAR_CONST = 2;
  const int INVOKE_FUNC = 1;
  const int INVOKE_PROPERTYGET = 2;
  const int INVOKE_PROPERTYPUT = 4;
  const int INVOKE_PROPERTYPUTREF = 8;
  const int TKIND_ENUM = 0;
  const int TKIND_MODULE = 2;
  const int TKIND_INTERFACE = 3;
  const int TKIND_DISPATCH = 4;
  const int TKIND_COCLASS = 5;
  const int TKIND_ALIAS = 6;
  const int VT_PTR = 26;
  const int VT_SAFEARRAY = 27;
  const int VT_CARRAY = 28;
  const int VT_USERDEFINED = 29;
  const int VT_HRESULT = 25;
  const int VT_VOID = 24;

  [StructLayout(LayoutKind.Sequential)]
  struct NativeVARDESC {
    public int memid;
    int _pad0;
    public IntPtr lpstrSchema;
    public IntPtr lpvarValue;
    public CT.ELEMDESC elemdescVar;
    public short wVarFlags;
    short _pad1;
    public int varkind;
  }

  class TypeRec {
    public string name;
    public string library;
    public string kind;
    public int typekind;
    public CT.ITypeInfo info;
    public List<Dictionary<string, object>> members = new List<Dictionary<string, object>>();
    public List<string> defaultImpls = new List<string>();
    public bool appObject;
  }

  static readonly HashSet<string> SkipMembers = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
    "QueryInterface", "AddRef", "Release", "GetTypeInfoCount", "GetTypeInfo", "GetIDsOfNames", "Invoke"
  };

  public static string DumpAll(string vbaPath, string officePath, string excelPath) {
    var types = new Dictionary<string, TypeRec>(StringComparer.OrdinalIgnoreCase);
    var globals = new Dictionary<string, Dictionary<string, object>>(StringComparer.OrdinalIgnoreCase);

    LoadLib(vbaPath, "VBA", types, globals, true);
    LoadLib(officePath, "Office", types, globals, false);
    LoadLib(excelPath, "Excel", types, globals, false);

    MergeCoclasses(types);

    var typesOut = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
    foreach (var kv in types) {
      var rec = kv.Value;
      if (rec.members.Count == 0 && rec.kind != "enum") continue;
      if (ShouldDropType(rec.name)) continue;
      var td = new Dictionary<string, object>();
      td["name"] = rec.name;
      td["library"] = rec.library;
      td["kind"] = rec.kind;
      td["members"] = rec.members;
      typesOut[rec.name] = td;
    }

    var globalsOut = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
    foreach (var kv in globals) globalsOut[kv.Value["name"].ToString()] = kv.Value;

    var root = new Dictionary<string, object>();
    root["types"] = typesOut;
    root["globals"] = globalsOut;
    return ToJson(root);
  }

  static bool ShouldDropType(string name) {
    if (string.IsNullOrEmpty(name)) return true;
    if (name == "IUnknown" || name == "IDispatch" || name == "IInspectable") return true;
    if (name.StartsWith("_") && !name.Equals("_Global", StringComparison.OrdinalIgnoreCase)) return true;
    return false;
  }

  static void LoadLib(string path, string library, Dictionary<string, TypeRec> types, Dictionary<string, Dictionary<string, object>> globals, bool allModulesGlobal) {
    if (string.IsNullOrEmpty(path)) return;
    CT.ITypeLib tlb;
    LoadTypeLibEx(path, 2, out tlb);
    int count = tlb.GetTypeInfoCount();
    var local = new List<TypeRec>();
    for (int i = 0; i < count; i++) {
      CT.ITypeInfo ti;
      tlb.GetTypeInfo(i, out ti);
      string name, doc, help;
      int ctx;
      ti.GetDocumentation(-1, out name, out doc, out ctx, out help);
      if (string.IsNullOrEmpty(name)) continue;
      if (name == "IUnknown" || name == "IDispatch") continue;

      IntPtr pAttr;
      ti.GetTypeAttr(out pAttr);
      CT.TYPEATTR attr = (CT.TYPEATTR)Marshal.PtrToStructure(pAttr, typeof(CT.TYPEATTR));
      int tkind = (int)attr.typekind;
      short cFuncs = attr.cFuncs;
      short cVars = attr.cVars;
      short cImpl = attr.cImplTypes;
      short flags = (short)attr.wTypeFlags;
      ti.ReleaseTypeAttr(pAttr);

      if (tkind == TKIND_ALIAS) continue;

      var rec = new TypeRec();
      rec.name = name;
      rec.library = library;
      rec.typekind = tkind;
      rec.info = ti;
      rec.appObject = (flags & 1) != 0;
      if (tkind == TKIND_ENUM) rec.kind = "enum";
      else if (tkind == TKIND_MODULE) rec.kind = "module";
      else rec.kind = "class";

      if (tkind == TKIND_COCLASS) {
        for (int j = 0; j < cImpl; j++) {
          CT.IMPLTYPEFLAGS implFlags;
          ti.GetImplTypeFlags(j, out implFlags);
          if ((((int)implFlags) & IMPLTYPEFLAG_FSOURCE) != 0) continue;
          int href;
          ti.GetRefTypeOfImplType(j, out href);
          CT.ITypeInfo impl;
          ti.GetRefTypeInfo(href, out impl);
          string iname, idoc, ihelp;
          int ictx;
          impl.GetDocumentation(-1, out iname, out idoc, out ictx, out ihelp);
            if ((((int)implFlags) & IMPLTYPEFLAG_FDEFAULT) != 0 || rec.defaultImpls.Count == 0) {
            if ((((int)implFlags) & IMPLTYPEFLAG_FDEFAULT) != 0) rec.defaultImpls.Insert(0, iname);
            else rec.defaultImpls.Add(iname);
          }
        }
      } else {
        rec.members.AddRange(ReadMembers(ti, cFuncs, cVars, tkind == TKIND_ENUM));
        if (library == "VBA" && rec.kind == "module") AliasVbaMembers(rec);
      }

      local.Add(rec);
      types[name] = rec;
    }

    foreach (var rec in local) {
      bool promote = allModulesGlobal && rec.kind == "module";
      if (rec.name.Equals("_Global", StringComparison.OrdinalIgnoreCase) || rec.name.Equals("Global", StringComparison.OrdinalIgnoreCase)) promote = true;
      if (rec.appObject) {
        AddGlobal(globals, rec.name, library, rec.name, null);
      }
      if (promote || rec.kind == "enum") {
        foreach (var mem in rec.members) {
          string mname = mem["name"] as string;
          if (!string.IsNullOrEmpty(mname) && mname.StartsWith("_B_", StringComparison.OrdinalIgnoreCase)) continue;
          string typeName = null;
          if (mem.ContainsKey("returnType")) {
            string rt = mem["returnType"] as string;
            if (!string.IsNullOrEmpty(rt) && rt != "Object" && rt != "Variant" && rt != "Any") typeName = rt;
          }
          if (rec.kind == "enum") typeName = rec.name;
          string owner = rec.kind == "module" || rec.name.StartsWith("_") ? library : rec.name;
          AddGlobal(globals, mname, owner, typeName, mem);
        }
      }
    }
  }

  static void MergeCoclasses(Dictionary<string, TypeRec> types) {
    foreach (var rec in types.Values) {
      if (rec.typekind != TKIND_COCLASS) continue;
      var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
      foreach (var implName in rec.defaultImpls) {
        TypeRec impl;
        if (!types.TryGetValue(implName, out impl)) continue;
        foreach (var mem in impl.members) {
          string key = mem["name"] + "\t" + mem["kind"];
          if (!seen.Add(key)) continue;
          rec.members.Add(mem);
        }
      }
      string hidden = "_" + rec.name;
      TypeRec hiddenRec;
      if (types.TryGetValue(hidden, out hiddenRec)) {
        foreach (var mem in hiddenRec.members) {
          string key = mem["name"] + "\t" + mem["kind"];
          if (!seen.Add(key)) continue;
          rec.members.Add(mem);
        }
      }
    }
  }

  static void AliasVbaMembers(TypeRec rec) {
    var extra = new List<Dictionary<string, object>>();
    var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    foreach (var mem in rec.members) {
      string n = mem["name"] as string;
      if (!string.IsNullOrEmpty(n)) seen.Add(n + "\t" + mem["kind"]);
    }
    foreach (var mem in rec.members) {
      string n = mem["name"] as string;
      string alias = VbaAlias(n);
      if (alias == null) continue;
      string key = alias + "\t" + mem["kind"];
      if (!seen.Add(key)) continue;
      extra.Add(CloneMember(mem, alias));
    }
    rec.members.AddRange(extra);
  }

  static string VbaAlias(string name) {
    if (string.IsNullOrEmpty(name)) return null;
    if (name.StartsWith("_B_str_", StringComparison.OrdinalIgnoreCase)) return name.Substring(7) + "$";
    if (name.StartsWith("_B_var_", StringComparison.OrdinalIgnoreCase)) return name.Substring(7);
    return null;
  }

  static Dictionary<string, object> CloneMember(Dictionary<string, object> mem, string newName) {
    var copy = new Dictionary<string, object>(mem);
    string old = mem["name"] as string;
    copy["name"] = newName;
    if (mem.ContainsKey("signature") && !string.IsNullOrEmpty(old)) {
      string sig = mem["signature"] as string;
      int idx = sig.IndexOf(old, StringComparison.Ordinal);
      if (idx >= 0) copy["signature"] = sig.Substring(0, idx) + newName + sig.Substring(idx + old.Length);
    }
    return copy;
  }

  static void AddGlobal(Dictionary<string, Dictionary<string, object>> globals, string name, string owner, string typeName, Dictionary<string, object> member) {
    if (string.IsNullOrEmpty(name) || SkipMembers.Contains(name)) return;
    var g = new Dictionary<string, object>();
    g["name"] = name;
    g["owner"] = owner;
    if (!string.IsNullOrEmpty(typeName)) g["typeName"] = typeName;
    if (member != null) g["member"] = member;
    globals[name] = g;
  }

  static List<Dictionary<string, object>> ReadMembers(CT.ITypeInfo ti, int cFuncs, int cVars, bool isEnum) {
    var list = new List<Dictionary<string, object>>();
    var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    for (int i = 0; i < cFuncs; i++) {
      IntPtr pfd;
      ti.GetFuncDesc(i, out pfd);
      try {
        CT.FUNCDESC fd = (CT.FUNCDESC)Marshal.PtrToStructure(pfd, typeof(CT.FUNCDESC));
        if ((((int)fd.wFuncFlags) & FUNCFLAG_FRESTRICTED) != 0) continue;
        var mem = ReadFunc(ti, fd);
        if (mem == null) continue;
        string key = mem["name"] + "\t" + mem["kind"];
        if (!seen.Add(key)) continue;
        list.Add(mem);
      } finally {
        ti.ReleaseFuncDesc(pfd);
      }
    }
    for (int i = 0; i < cVars; i++) {
      IntPtr pvd;
      ti.GetVarDesc(i, out pvd);
      try {
        NativeVARDESC vd = (NativeVARDESC)Marshal.PtrToStructure(pvd, typeof(NativeVARDESC));
        string[] names = new string[1];
        int n;
        ti.GetNames(vd.memid, names, 1, out n);
        string vname = n > 0 ? names[0] : "";
        if (string.IsNullOrEmpty(vname) || SkipMembers.Contains(vname)) continue;
        string summary, help;
        int ctx;
        string dummy;
        ti.GetDocumentation(vd.memid, out dummy, out summary, out ctx, out help);
        var mem = new Dictionary<string, object>();
        mem["name"] = vname;
        if (isEnum || vd.varkind == VAR_CONST) {
          mem["kind"] = "Const";
          string val = "";
          if (vd.lpvarValue != IntPtr.Zero) val = FormatVariant(vd.lpvarValue);
          mem["signature"] = "Const " + vname + (val.Length > 0 ? " = " + val : "");
          if (val.Length > 0) mem["defaultValue"] = val;
          mem["returnType"] = TypeName(ti, vd.elemdescVar.tdesc);
        } else {
          mem["kind"] = "Property Get";
          string rt = TypeName(ti, vd.elemdescVar.tdesc);
          mem["signature"] = "Property Get " + vname + "() As " + rt;
          mem["returnType"] = rt;
          mem["params"] = new List<object>();
        }
        if (!string.IsNullOrEmpty(summary)) mem["summary"] = summary.Trim();
        string key = vname + "\t" + mem["kind"];
        if (!seen.Add(key)) continue;
        list.Add(mem);
      } finally {
        ti.ReleaseVarDesc(pvd);
      }
    }
    return list;
  }

  static Dictionary<string, object> ReadFunc(CT.ITypeInfo ti, CT.FUNCDESC fd) {
    int cParams = fd.cParams;
    string[] names = new string[Math.Max(1, cParams + 1)];
    int n;
    ti.GetNames(fd.memid, names, names.Length, out n);
    string fname = n > 0 ? names[0] : "";
    if (string.IsNullOrEmpty(fname) || SkipMembers.Contains(fname)) return null;

    string summary, help;
    int ctx;
    string dummy;
    ti.GetDocumentation(fd.memid, out dummy, out summary, out ctx, out help);

    int inv = (int)fd.invkind;
    var rawParams = new List<Dictionary<string, object>>();
    int elemSize = Marshal.SizeOf(typeof(CT.ELEMDESC));
    string retType = TypeName(ti, fd.elemdescFunc.tdesc);
    bool retIsVoid = retType == "" || retType == "Void" || retType == "HRESULT";

    for (int i = 0; i < cParams; i++) {
      CT.ELEMDESC ed = (CT.ELEMDESC)Marshal.PtrToStructure(
        new IntPtr(fd.lprgelemdescParam.ToInt64() + i * elemSize),
        typeof(CT.ELEMDESC));
      int pflags = (int)ed.desc.paramdesc.wParamFlags;
      if ((pflags & PARAMFLAG_FLCID) != 0) continue;
      string pname = (i + 1 < n && !string.IsNullOrEmpty(names[i + 1])) ? names[i + 1] : "arg" + (i + 1);
      string ptype = TypeName(ti, ed.tdesc);
      if ((pflags & PARAMFLAG_FRETVAL) != 0) {
        retType = ptype;
        retIsVoid = false;
        continue;
      }
      var p = new Dictionary<string, object>();
      p["name"] = pname;
      if (!string.IsNullOrEmpty(ptype) && ptype != "Variant") p["type"] = ptype;
      if ((pflags & PARAMFLAG_FOPT) != 0) p["optional"] = true;
      if ((pflags & PARAMFLAG_FHASDEFAULT) != 0) {
        IntPtr pex = ed.desc.paramdesc.lpVarValue;
        if (pex != IntPtr.Zero) {
          int offset = IntPtr.Size == 8 ? 8 : 4;
          string def = FormatVariant(new IntPtr(pex.ToInt64() + offset));
          if (def.Length > 0) p["defaultValue"] = def;
        }
      }
      rawParams.Add(p);
    }

    string kind;
    if (inv == INVOKE_PROPERTYGET) kind = "Property Get";
    else if (inv == INVOKE_PROPERTYPUT) kind = "Property Let";
    else if (inv == INVOKE_PROPERTYPUTREF) kind = "Property Set";
    else if (retIsVoid) kind = "Sub";
    else kind = "Function";

    if (kind == "Sub") retType = "";
    if (retType == "HRESULT") retType = "";

    var mem = new Dictionary<string, object>();
    mem["name"] = fname;
    mem["kind"] = kind;
    mem["signature"] = BuildSignature(kind, fname, rawParams, retType);
    if (rawParams.Count > 0) mem["params"] = ToObjectList(rawParams);
    if (!string.IsNullOrEmpty(retType)) mem["returnType"] = retType;
    if (!string.IsNullOrEmpty(summary)) mem["summary"] = summary.Trim();
    return mem;
  }

  static List<object> ToObjectList(List<Dictionary<string, object>> items) {
    var list = new List<object>();
    foreach (var x in items) list.Add(x);
    return list;
  }

  static string BuildSignature(string kind, string name, List<Dictionary<string, object>> pars, string retType) {
    var sb = new StringBuilder();
    if (kind.StartsWith("Property")) sb.Append(kind);
    else sb.Append(kind);
    sb.Append(' ').Append(name);
    sb.Append('(');
    for (int i = 0; i < pars.Count; i++) {
      if (i > 0) sb.Append(", ");
      var p = pars[i];
      bool opt = p.ContainsKey("optional") && (bool)p["optional"];
      if (opt) sb.Append("Optional ");
      sb.Append(p["name"]);
      if (p.ContainsKey("type")) sb.Append(" As ").Append(p["type"]);
      if (p.ContainsKey("defaultValue")) sb.Append(" = ").Append(p["defaultValue"]);
    }
    sb.Append(')');
    if (!string.IsNullOrEmpty(retType) && kind != "Sub") sb.Append(" As ").Append(retType);
    return sb.ToString();
  }

  static string TypeName(CT.ITypeInfo ti, CT.TYPEDESC td) {
    int vt = td.vt & 0x0FFF;
    if (vt == VT_PTR || vt == VT_SAFEARRAY || vt == VT_CARRAY) {
      if (td.lpValue == IntPtr.Zero) return vt == VT_PTR ? "Variant" : "Variant()";
      CT.TYPEDESC inner = (CT.TYPEDESC)Marshal.PtrToStructure(td.lpValue, typeof(CT.TYPEDESC));
      string innerName = TypeName(ti, inner);
      if (vt == VT_PTR) return innerName;
      return string.IsNullOrEmpty(innerName) ? "Variant()" : innerName + "()";
    }
    if (vt == VT_USERDEFINED) {
      try {
        int href = (int)(td.lpValue.ToInt64() & 0xFFFFFFFF);
        CT.ITypeInfo refInfo;
        ti.GetRefTypeInfo(href, out refInfo);
        string name, doc, help;
        int ctx;
        refInfo.GetDocumentation(-1, out name, out doc, out ctx, out help);
        IntPtr pAttr;
        refInfo.GetTypeAttr(out pAttr);
        CT.TYPEATTR attr = (CT.TYPEATTR)Marshal.PtrToStructure(pAttr, typeof(CT.TYPEATTR));
        int tkind = (int)attr.typekind;
        CT.TYPEDESC aliasTd = attr.tdescAlias;
        refInfo.ReleaseTypeAttr(pAttr);
        if (tkind == TKIND_ALIAS) {
          string aliased = TypeName(refInfo, aliasTd);
          return string.IsNullOrEmpty(name) ? aliased : name;
        }
        if (!string.IsNullOrEmpty(name) && name.StartsWith("_") && name.Length > 1) return name.Substring(1);
        return name ?? "";
      } catch {
        return "Object";
      }
    }
    return VarToVba(vt);
  }

  static string VarToVba(int vt) {
    switch (vt) {
      case 0: return "";
      case 1: return "Null";
      case 2: return "Integer";
      case 3: return "Long";
      case 4: return "Single";
      case 5: return "Double";
      case 6: return "Currency";
      case 7: return "Date";
      case 8: return "String";
      case 9: return "Object";
      case 11: return "Boolean";
      case 12: return "Variant";
      case 13: return "Object";
      case 14: return "Decimal";
      case 16: return "Integer";
      case 17: return "Byte";
      case 18: return "Integer";
      case 19: return "Long";
      case 20: return "LongLong";
      case 21: return "LongLong";
      case 22: return "Long";
      case 23: return "Long";
      case 24: return "Void";
      case 25: return "HRESULT";
      case 30: return "String";
      case 31: return "String";
      default: return "Variant";
    }
  }

  static string FormatVariant(IntPtr pVar) {
    if (pVar == IntPtr.Zero) return "";
    try {
      object o = Marshal.GetObjectForNativeVariant(pVar);
      if (o == null) return "Null";
      if (o is string) return "\"" + ((string)o).Replace("\"", "\"\"") + "\"";
      if (o is bool) return (bool)o ? "True" : "False";
      if (o is DateTime) return ((DateTime)o).ToString("d", CultureInfo.InvariantCulture);
      if (o is double || o is float) return Convert.ToString(o, CultureInfo.InvariantCulture);
      return Convert.ToString(o, CultureInfo.InvariantCulture);
    } catch {
      return "";
    }
  }

  static string ToJson(object o) {
    var sb = new StringBuilder();
    WriteJson(sb, o);
    return sb.ToString();
  }

  static void WriteJson(StringBuilder sb, object o) {
    if (o == null) { sb.Append("null"); return; }
    string s = o as string;
    if (s != null) { sb.Append('"').Append(Esc(s)).Append('"'); return; }
    if (o is bool) { sb.Append((bool)o ? "true" : "false"); return; }
    if (o is int || o is long || o is short || o is byte) { sb.Append(o.ToString()); return; }
    if (o is double || o is float || o is decimal) { sb.Append(Convert.ToString(o, CultureInfo.InvariantCulture)); return; }
    IDictionary dict = o as IDictionary;
    if (dict != null) {
      sb.Append('{');
      bool first = true;
      foreach (DictionaryEntry kv in dict) {
        if (!first) sb.Append(',');
        first = false;
        sb.Append('"').Append(Esc(kv.Key.ToString())).Append('"').Append(':');
        WriteJson(sb, kv.Value);
      }
      sb.Append('}');
      return;
    }
    IEnumerable en = o as IEnumerable;
    if (en != null) {
      sb.Append('[');
      bool first = true;
      foreach (object item in en) {
        if (!first) sb.Append(',');
        first = false;
        WriteJson(sb, item);
      }
      sb.Append(']');
      return;
    }
    sb.Append('"').Append(Esc(Convert.ToString(o, CultureInfo.InvariantCulture))).Append('"');
  }

  static string Esc(string s) {
    if (string.IsNullOrEmpty(s)) return "";
    var sb = new StringBuilder();
    foreach (char c in s) {
      if (c == '"') sb.Append("\\\"");
      else if (c == '\\') sb.Append("\\\\");
      else if (c == '\n') sb.Append("\\n");
      else if (c == '\r') sb.Append("\\r");
      else if (c == '\t') sb.Append("\\t");
      else if (c < 32) sb.Append("\\u").Append(((int)c).ToString("x4"));
      else sb.Append(c);
    }
    return sb.ToString();
  }
}
'@

Add-Type -TypeDefinition $cs -Language CSharp

function Find-FirstExisting([string[]]$paths) {
  foreach ($p in $paths) { if ($p -and (Test-Path -LiteralPath $p)) { return $p } }
  return $null
}

$vba = Find-FirstExisting @(
  "$env:ProgramFiles\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\VBA\VBA7.1\VBE7.DLL",
  "$env:ProgramFiles\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL"
)
$office = Find-FirstExisting @(
  "$env:ProgramFiles\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\OFFICE16\MSO.DLL",
  "$env:ProgramFiles\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\OFFICE16\MSO.DLL",
  "$env:ProgramFiles\Common Files\Microsoft Shared\OFFICE16\MSO.DLL"
)
$excel = Find-FirstExisting @(
  "$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE",
  "$env:ProgramFiles\Microsoft Office\Office16\EXCEL.EXE"
)

if (-not $vba) { throw 'VBE7.DLL not found' }
if (-not $excel) { throw 'EXCEL.EXE not found' }
if (-not $office) { Write-Warning 'MSO.DLL not found; dumping VBA + Excel only' }

Write-Host "VBA    $vba"
Write-Host "Office $office"
Write-Host "Excel  $excel"

$outDir = Split-Path -Parent $outFile
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

$json = [TypeLibDump]::DumpAll($vba, $office, $excel)
[System.IO.File]::WriteAllText($outFile, $json)

Write-Host ("Wrote {0} ({1:N1} KB)" -f $outFile, ((Get-Item $outFile).Length / 1KB))
