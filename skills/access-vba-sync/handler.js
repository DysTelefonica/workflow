"use strict";

const fs = require("fs");
const fsp = fs.promises;
const path = require("path");
const { spawn } = require("child_process");

function uniq(arr) {
  const out = [];
  const seen = new Set();
  for (const v of arr || []) {
    const k = String(v);
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(k);
  }
  return out;
}

function powershellExe() {
  return process.env.POWERSHELL_EXE || "powershell.exe";
}

function isAccessDbFileName(name) {
  const ext = path.extname(name).toLowerCase();
  return ext === ".accdb" || ext === ".accde" || ext === ".mdb" || ext === ".mde";
}

function isWatchedExt(filePath) {
  const name = path.basename(filePath).toLowerCase();
  if (name.endsWith(".form.txt")) return true;
  const ext = path.extname(filePath).toLowerCase();
  return ext === ".bas" || ext === ".cls" || ext === ".frm";
}

function moduleNameFromFile(filePath) {
  const base = path.basename(filePath);
  if (base.toLowerCase().endsWith(".form.txt")) {
    return base.slice(0, -".form.txt".length);
  }
  return path.basename(filePath, path.extname(filePath));
}

function logicalModuleNameVariants(moduleName) {
  const raw = String(moduleName || "").trim();
  const base = raw.replace(/^(Form|Report)_/i, "");
  const out = [];
  const push = (value) => {
    const v = String(value || "").trim();
    if (!v || out.includes(v)) return;
    out.push(v);
  };

  push(raw);
  push(base);
  push(`Form_${base}`);
  push(`Report_${base}`);
  return out;
}

function isVbaMetadataLine(line) {
  const trim = String(line || "").trim();
  if (!trim) return false;
  return (
    /^VERSION\s+\d+(\.\d+)?\s+CLASS$/i.test(trim) ||
    /^BEGIN\b/i.test(trim) ||
    /^END$/i.test(trim) ||
    /^(MultiUse|Persistable|DataBindingBehavior|DataSourceBehavior|MTSTransactionMode)\s*=/i.test(trim) ||
    /^Attribute\s+VB_/i.test(trim)
  );
}

function isVbaOptionDirectiveLine(line) {
  const trim = String(line || "").trim();
  return /^Option\s+(Compare\s+\w+|Explicit|Base\s+\d+|Private\s+Module)$/i.test(trim);
}

function sanitizeVbaImportText(text) {
  const normalized = String(text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = normalized.split("\n");
  if (lines.length > 0 && lines[0].charCodeAt(0) === 0xfeff) {
    lines[0] = lines[0].slice(1);
  }

  let start = 0;
  while (start < lines.length) {
    const trim = lines[start].trim();
    if (!trim || isVbaMetadataLine(lines[start])) {
      start++;
      continue;
    }
    break;
  }

  const out = [];
  const seenOptions = new Set();
  let inDirectiveBlock = true;

  for (let i = start; i < lines.length; i++) {
    const line = lines[i];
    const trim = line.trim();

    if (inDirectiveBlock) {
      if (!trim) {
        out.push(line);
        continue;
      }

      if (isVbaMetadataLine(line)) continue;

      if (isVbaOptionDirectiveLine(line)) {
        const key = trim.toLowerCase();
        if (!seenOptions.has(key)) {
          seenOptions.add(key);
          out.push(line);
        }
        continue;
      }

      inDirectiveBlock = false;
    }

    out.push(line);
  }

  while (out.length > 0 && !out[0].trim()) out.shift();
  return out.join("\r\n");
}

function preferredNewline(text) {
  return String(text || "").includes("\r\n") ? "\r\n" : "\n";
}

function normalizeNewlines(text, newline = "\n") {
  return String(text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n/g, newline);
}

function normalizePathForComparison(targetPath) {
  const resolved = path.resolve(String(targetPath || ""));
  const trimmed = resolved.replace(/[\\\/]+$/, "");
  return process.platform === "win32" ? trimmed.toLowerCase() : trimmed;
}

function splitCodeBehindSection(text) {
  const normalized = normalizeNewlines(text, "\n");
  const match = /^([ \t]*CodeBehind\w*[^\r\n]*)(?:\n|$)/im.exec(normalized);
  if (!match || match.index == null) return null;

  const start = match.index;
  const markerLine = match[1];
  const markerEnd = start + match[0].length;

  return {
    before: normalized.slice(0, start),
    markerLine,
    body: normalized.slice(markerEnd)
  };
}

function splitVbaMetadataHeaderText(text) {
  const normalized = normalizeNewlines(text, "\n");
  const lines = normalized.split("\n");
  if (lines.length > 0 && lines[0].charCodeAt(0) === 0xfeff) {
    lines[0] = lines[0].slice(1);
  }

  const header = [];
  let index = 0;
  while (index < lines.length) {
    const line = lines[index];
    const trim = String(line || "").trim();
    if (!trim || isVbaMetadataLine(line) || isVbaOptionDirectiveLine(line)) {
      header.push(line);
      index += 1;
      continue;
    }
    break;
  }

  while (header.length > 0 && !String(header[header.length - 1] || "").trim()) {
    header.pop();
  }

  let bodyLines = lines.slice(index);
  while (bodyLines.length > 0 && !String(bodyLines[0] || "").trim()) {
    bodyLines.shift();
  }

  return {
    header: header.join("\n"),
    body: bodyLines.join("\n")
  };
}

function joinCodeBehindBodyText(header, body, newline) {
  const parts = [];
  const normalizedHeader = normalizeNewlines(header, "\n").replace(/\n+$/g, "");
  const normalizedBody = normalizeNewlines(body, "\n").replace(/^\n+/g, "");

  if (normalizedHeader) parts.push(normalizedHeader);
  if (normalizedBody) parts.push(normalizedBody);

  return parts.join("\n").replace(/\n/g, newline);
}

function mergeDocumentCodeBehindText(documentText, clsText) {
  const newline = preferredNewline(documentText) || "\r\n";
  const section = splitCodeBehindSection(documentText);
  if (!section) {
    throw new Error("No se encontró ningún marcador CodeBehind* en el documento.");
  }

  const sanitizedCls = sanitizeVbaImportText(clsText);
  const documentCode = splitVbaMetadataHeaderText(section.body);
  const clsCode = splitVbaMetadataHeaderText(sanitizedCls);
  const effectiveHeader = documentCode.header || clsCode.header;
  const mergedBody = joinCodeBehindBodyText(effectiveHeader, clsCode.body, newline);
  const normalizedBefore = normalizeNewlines(section.before, newline);

  return normalizedBefore + section.markerLine + newline + mergedBody;
}

function countMeaningfulVbaBodyLines(text) {
  const parts = splitVbaMetadataHeaderText(sanitizeVbaImportText(text || ""));
  const bodyLines = String(parts.body || "")
    .split(/\r?\n/)
    .map((line) => String(line || "").trim())
    .filter((line) => line && !line.startsWith("'"));
  return bodyLines.length;
}

function normalizedMeaningfulVbaBody(text) {
  const parts = splitVbaMetadataHeaderText(sanitizeVbaImportText(text || ""));
  return String(parts.body || "")
    .split(/\r?\n/)
    .map((line) => String(line || "").trim())
    .filter((line) => line && !line.startsWith("'"))
    .join("\n");
}

function pickBestDocumentClsPath({ canonicalPath, legacyPath }) {
  const existing = uniq([canonicalPath, legacyPath].filter(Boolean)).filter((candidate) => fs.existsSync(candidate));
  if (existing.length === 0) return { clsPath: null, mirrorPaths: [], warning: null, error: null };

  const scored = existing.map((candidate) => {
    let content = "";
    let mtimeMs = 0;
    try {
      content = fs.readFileSync(candidate, "utf8");
    } catch {}
    try {
      mtimeMs = fs.statSync(candidate).mtimeMs || 0;
    } catch {}
    return {
      path: candidate,
      meaningfulLines: countMeaningfulVbaBodyLines(content),
      normalizedBody: normalizedMeaningfulVbaBody(content),
      mtimeMs,
      isCanonical: candidate === canonicalPath
    };
  });

  const canonical = scored.find((item) => item.path === canonicalPath) || null;
  const legacy = scored.find((item) => item.path === legacyPath) || null;
  const newest = [...scored].sort((a, b) => b.mtimeMs - a.mtimeMs)[0];
  const meaningful = scored.filter((item) => item.meaningfulLines > 0);

  if (scored.length === 1) {
    return { clsPath: scored[0].path, mirrorPaths: [], warning: null, error: null };
  }

  if (meaningful.length === 0) {
    const preferred = canonical || newest;
    return {
      clsPath: preferred.path,
      mirrorPaths: scored.filter((item) => item.path !== preferred.path).map((item) => item.path),
      warning: `⚠️  Ambos sidecars .cls están vacíos o casi vacíos. Se usó ${path.basename(preferred.path)} por criterio canónico/reciente y se unificarán los gemelos.`,
      error: null
    };
  }

  if (meaningful.length === 1) {
    const preferred = meaningful[0];
    return {
      clsPath: preferred.path,
      mirrorPaths: scored.filter((item) => item.path !== preferred.path).map((item) => item.path),
      warning: `⚠️  Se eligió ${path.basename(preferred.path)} porque el otro sidecar está vacío o casi vacío. Se unificarán ambos .cls.`,
      error: null
    };
  }

  const distinctBodies = new Set(meaningful.map((item) => item.normalizedBody));
  if (distinctBodies.size === 1) {
    const preferred = canonical || newest;
    return {
      clsPath: preferred.path,
      mirrorPaths: scored.filter((item) => item.path !== preferred.path).map((item) => item.path),
      warning: null,
      error: null
    };
  }

  const preferred = newest;
  return {
    clsPath: preferred.path,
    mirrorPaths: scored.filter((item) => item.path !== preferred.path).map((item) => item.path),
    warning: `⚠️  Se detectaron sidecars .cls divergentes (${meaningful.map((item) => path.basename(item.path)).join(", ")}). Se priorizó ${path.basename(preferred.path)} por ser el modificado más recientemente y se unificarán los gemelos antes del import.`,
    error: null
  };
}

function unifiedDiff(aLines, bLines, labelA, labelB, contextLines = 3) {
  const max = Math.max(aLines.length, bLines.length);
  let hasChanges = false;
  const rowsByLine = [];
  const changedLineIndexes = new Set();

  for (let i = 0; i < max; i++) {
    const a = i < aLines.length ? aLines[i] : undefined;
    const b = i < bLines.length ? bLines[i] : undefined;
    if (a === b) {
      rowsByLine.push([` ${a}`]);
      continue;
    }
    hasChanges = true;
    changedLineIndexes.add(i);
    const rows = [];
    if (a !== undefined) rows.push(`-${a}`);
    if (b !== undefined) rows.push(`+${b}`);
    rowsByLine.push(rows);
  }

  if (!hasChanges) return null;

  const context = Number.isFinite(contextLines) ? Math.max(0, Number(contextLines)) : max;
  const rows = [];
  let omitted = false;
  for (let i = 0; i < rowsByLine.length; i++) {
    let include = changedLineIndexes.has(i);
    if (!include) {
      for (const changedIndex of changedLineIndexes) {
        if (Math.abs(i - changedIndex) <= context) {
          include = true;
          break;
        }
      }
    }

    if (!include) {
      if (!omitted) {
        rows.push(" ...");
        omitted = true;
      }
      continue;
    }

    omitted = false;
    rows.push(...rowsByLine[i]);
  }

  return [
    `--- ${labelA}`,
    `+++ ${labelB}`,
    `@@ -1,${aLines.length} +1,${bLines.length} @@`,
    ...rows
  ].join("\n");
}

function parseModuleResults(stdout) {
  const line = String(stdout || "").split(/\r?\n/).find((l) => l.startsWith("##MODULE_RESULTS:"));
  if (!line) return null;
  try { return JSON.parse(line.slice("##MODULE_RESULTS:".length)); } catch { return null; }
}

function extractJsonTextAt(source, start) {
  const opening = source[start];
  const closingFor = { "{": "}", "[": "]" };
  if (!closingFor[opening]) return null;

  const stack = [closingFor[opening]];
  let inString = false;
  let escaped = false;

  for (let i = start + 1; i < source.length; i++) {
    const ch = source[i];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (ch === "\\") {
      escaped = inString;
      continue;
    }
    if (ch === "\"") {
      inString = !inString;
      continue;
    }
    if (inString) continue;

    if (ch === "{" || ch === "[") {
      stack.push(closingFor[ch]);
      continue;
    }
    if (ch === "}" || ch === "]") {
      if (stack.pop() !== ch) return null;
      if (stack.length === 0) return source.slice(start, i + 1);
    }
  }

  return null;
}

function extractFirstJsonText(text) {
  const source = String(text || "");
  for (let start = 0; start < source.length; start++) {
    if (source[start] !== "{" && source[start] !== "[") continue;
    const jsonText = extractJsonTextAt(source, start);
    if (!jsonText) continue;
    try {
      JSON.parse(jsonText);
      return jsonText;
    } catch {}
  }

  return null;
}

function parseJsonFromStdout(stdout, label = "salida JSON") {
  const text = String(stdout || "").trim();
  if (!text) throw new Error(`No se recibió ${label}.`);

  try {
    return JSON.parse(text);
  } catch {}

  const jsonText = extractFirstJsonText(text);
  if (jsonText) {
    try {
      return JSON.parse(jsonText);
    } catch {}
  }

  throw new Error(`No se pudo parsear ${label}. La salida de PowerShell no contiene JSON válido.`);
}

function parseVbaArgsJson(argsJson) {
  if (argsJson == null || argsJson === "") return [];
  let parsed;
  try {
    parsed = JSON.parse(String(argsJson));
  } catch (err) {
    throw new Error(`--args-json debe ser un array JSON válido: ${err && err.message ? err.message : String(err)}`);
  }
  if (!Array.isArray(parsed)) {
    throw new Error("--args-json debe ser un array JSON. Ejemplo: --args-json \"[123, \\\"texto\\\", true]\"");
  }
  for (const value of parsed) {
    const t = typeof value;
    if (!(value === null || t === "string" || t === "number" || t === "boolean")) {
      throw new Error("--args-json solo soporta valores simples: string, number, boolean o null.");
    }
  }
  return parsed;
}

function deepEqualJson(a, b) {
  return JSON.stringify(a) === JSON.stringify(b);
}

function getByPath(obj, dottedPath) {
  if (!dottedPath) return obj;
  let current = obj;
  for (const part of String(dottedPath).split(".")) {
    if (current == null || !(part in Object(current))) return undefined;
    current = current[part];
  }
  return current;
}

function objectContains(actual, expected) {
  if (expected == null || typeof expected !== "object" || Array.isArray(expected)) {
    return deepEqualJson(actual, expected);
  }
  if (actual == null || typeof actual !== "object") return false;
  for (const [key, expectedValue] of Object.entries(expected)) {
    if (!objectContains(actual[key], expectedValue)) return false;
  }
  return true;
}

function evaluateVbaTestExpectation(result, expect = {}) {
  const failures = [];
  const effectiveExpect = expect && typeof expect === "object" ? expect : {};

  if (Object.prototype.hasOwnProperty.call(effectiveExpect, "ok") && result.ok !== effectiveExpect.ok) {
    failures.push(`ok esperado ${effectiveExpect.ok}, recibido ${result.ok}`);
  } else if (!Object.prototype.hasOwnProperty.call(effectiveExpect, "ok") && result.ok !== true) {
    failures.push(`ok esperado true, recibido ${result.ok}`);
  }

  if (Object.prototype.hasOwnProperty.call(effectiveExpect, "returnValue") && !deepEqualJson(result.returnValue, effectiveExpect.returnValue)) {
    failures.push(`returnValue esperado ${JSON.stringify(effectiveExpect.returnValue)}, recibido ${JSON.stringify(result.returnValue)}`);
  }

  if (Object.prototype.hasOwnProperty.call(effectiveExpect, "value")) {
    const value = result && result.payload && Object.prototype.hasOwnProperty.call(result.payload, "value")
      ? result.payload.value
      : result.returnValue;
    if (!deepEqualJson(value, effectiveExpect.value)) {
      failures.push(`value esperado ${JSON.stringify(effectiveExpect.value)}, recibido ${JSON.stringify(value)}`);
    }
  }

  if (effectiveExpect.payloadContains && !objectContains(result.payload, effectiveExpect.payloadContains)) {
    failures.push(`payload no contiene ${JSON.stringify(effectiveExpect.payloadContains)}`);
  }

  if (effectiveExpect.errorContains) {
    const errorText = String(result.error || "");
    if (!errorText.includes(String(effectiveExpect.errorContains))) {
      failures.push(`error no contiene ${JSON.stringify(effectiveExpect.errorContains)}; recibido ${JSON.stringify(errorText)}`);
    }
  }

  if (effectiveExpect.pathEquals && typeof effectiveExpect.pathEquals === "object") {
    for (const [pathKey, expectedValue] of Object.entries(effectiveExpect.pathEquals)) {
      const actualValue = getByPath(result, pathKey);
      if (!deepEqualJson(actualValue, expectedValue)) {
        failures.push(`${pathKey} esperado ${JSON.stringify(expectedValue)}, recibido ${JSON.stringify(actualValue)}`);
      }
    }
  }

  return {
    ok: failures.length === 0,
    failures
  };
}

function logModuleResults(results) {
  if (!Array.isArray(results) || results.length === 0) return;
  const errors = results.filter((r) => r.status === "error");
  if (results.length === 1 && errors.length === 0) return;
  const label = errors.length > 0 ? `${results.length - errors.length} ok, ${errors.length} error(s)` : `${results.length} ok`;
  console.log(`\n📋 Resultado por módulo (${label}):`);
  for (const r of results) {
    if (r.status === "ok") {
      console.log(`   ✅ ${r.module}`);
    } else {
      console.log(`   ❌ ${r.module} — ${r.error || "error desconocido"}`);
    }
  }
}

async function copyDirRecursive(src, dest) {
  await fsp.mkdir(dest, { recursive: true });
  const entries = await fsp.readdir(src, { withFileTypes: true });
  for (const entry of entries) {
    const srcEntry = path.join(src, entry.name);
    const destEntry = path.join(dest, entry.name);
    if (entry.isDirectory()) {
      await copyDirRecursive(srcEntry, destEntry);
    } else {
      await fsp.copyFile(srcEntry, destEntry);
    }
  }
}

class AccessVbaSyncSkill {
  constructor(options = {}) {
    this.skillDir = options.skillDir || __dirname;
    this.projectRoot = options.projectRoot || process.cwd();
    this.destinationRoot = options.destinationRoot || "src";
    this.debounceMs = Number.isFinite(options.debounceMs) ? options.debounceMs : 600;
    this.autoExportOnEnd = options.autoExportOnEnd !== false;
    this.password = options.password || process.env.ACCESS_VBA_PASSWORD || null;
    this.passwordSource = options.password ? "cli" : (process.env.ACCESS_VBA_PASSWORD ? "env" : null);
    this.allowStartupExecution = options.allowStartupExecution === true;
    this._passwordArgWarningShown = false;

    this.vbaManagerPath = path.join(this.skillDir, "VBAManager.ps1");
    this.stateDir = path.join(this.projectRoot, ".access-vba-skill");
    this.stateFile = path.join(this.stateDir, "session.json");

    this.session = {
      active: false,
      startedAt: null,
      accessPath: null,
      destinationRoot: null,
      modulesPath: null,
      changedModules: [],
      lastSyncAt: null
    };

    this.watcher = null;
    this.pendingModules = new Set();
    this.debounceTimer = null;
    this.importing = false;
    this.fileSignatures = new Map();
    this.lastImportedAtByModule = new Map();
    this.cooldownMs = Number.isFinite(options.cooldownMs) ? options.cooldownMs : 2000;
    this._accessOpenDeferred = false;
  }

  async ensureReady() {
    await fsp.mkdir(this.stateDir, { recursive: true });
    await fsp.access(this.vbaManagerPath, fs.constants.F_OK);
  }

  async loadSessionFromDisk() {
    try {
      const raw = await fsp.readFile(this.stateFile, "utf8");
      const parsed = JSON.parse(raw);
      this.session = Object.assign({}, this.session, parsed);
    } catch (err) {
      if (err && err.code === "ENOENT") return;
      console.warn(`WARN: no se pudo cargar session.json (${this.stateFile}): ${err && err.message ? err.message : String(err)}`);
      return;
    }
  }

  async saveSessionToDisk() {
    await fsp.mkdir(this.stateDir, { recursive: true });
    await fsp.writeFile(this.stateFile, JSON.stringify(this.session, null, 2), "utf8");
  }

  async detectAccessPath({ accessPath } = {}) {
    if (accessPath) {
      const resolved = path.resolve(this.projectRoot, accessPath);
      const ext = path.extname(resolved).toLowerCase();
      if (!isAccessDbFileName(resolved)) {
        throw new Error(`--access debe ser .accdb/.accde/.mdb/.mde. Recibido: ${ext || "(sin extensión)"}`);
      }

      if (normalizePathForComparison(path.dirname(resolved)) !== normalizePathForComparison(this.projectRoot)) {
        throw new Error(`La BD debe estar en el directorio desde donde ejecutás el CLI (${this.projectRoot}). Recibido: ${resolved}`);
      }

      try {
        const st = await fsp.stat(resolved);
        if (!st.isFile()) throw new Error("no es archivo");
      } catch {
        throw new Error(`No existe el archivo indicado por --access: ${resolved}`);
      }

      return resolved;
    }

    const entries = await fsp.readdir(this.projectRoot, { withFileTypes: true });
    const candidates = entries
      .filter((e) => e.isFile() && isAccessDbFileName(e.name))
      .map((e) => e.name)
      .sort((a, b) => a.localeCompare(b, "es", { sensitivity: "base" }));

    if (candidates.length === 0) return null;

    if (candidates.length > 1) {
      console.log("⚠️  Se encontraron varias BDs en CWD; eligiendo determinista (alfabético):");
      for (const c of candidates) console.log(`   - ${c}`);
    }

    return path.join(this.projectRoot, candidates[0]);
  }

  resolveDestinationRoot() {
    return path.resolve(this.projectRoot, this.destinationRoot);
  }

  modulesPathFor(accessPath, destinationRootAbs) {
    return destinationRootAbs;
  }

  async initializeSession({ accessPath, performInitialExport = false } = {}) {
    await this.ensureReady();
    await this.loadSessionFromDisk();

    const detectedAccess = await this.detectAccessPath({ accessPath });
    if (!detectedAccess) {
      throw new Error("No se encontró ninguna BD .accdb/.accde/.mdb/.mde en CWD y no se pasó --access.");
    }

    const destinationRootAbs = this.resolveDestinationRoot();
    await fsp.mkdir(destinationRootAbs, { recursive: true });
    const modulesPath = this.modulesPathFor(detectedAccess, destinationRootAbs);

    this.session.active = true;
    this.session.startedAt = this.session.startedAt || new Date().toISOString();
    this.session.accessPath = detectedAccess;
    this.session.destinationRoot = destinationRootAbs;
    this.session.modulesPath = modulesPath;
    this.session.changedModules = Array.isArray(this.session.changedModules) ? this.session.changedModules : [];
    this.session.lastSyncAt = this.session.lastSyncAt || null;
    this.session.pendingModules = [];
    this.session.watcherPid = null;

    await this.saveSessionToDisk();

    if (performInitialExport) {
      await this.backupSrcIfExists();
      console.log("🚀 Export inicial (todos los módulos)...");
      await this.runVbaManager({
        action: "Export",
        accessPath: detectedAccess,
        destinationRootAbs
      });
      console.log("✓ Export inicial completado");
      console.log(`📁 modulesPath: ${modulesPath}`);
      await this.saveSessionToDisk();
      return;
    }

    console.log("ℹ️  Sesión preparada sin export inicial.");
    console.log(`📁 modulesPath: ${modulesPath}`);
  }

  async resolveCommandContext({ accessPath, initializeIfNeeded = false, allowSessionMutation = false } = {}) {
    await this.ensureReady();
    await this.loadSessionFromDisk();

    const requestedAccessPath = accessPath || null;
    if (!this.session.active && initializeIfNeeded) {
      await this.initializeSession({ accessPath: requestedAccessPath, performInitialExport: false });
      await this.loadSessionFromDisk();
    }

    const destinationRootAbs = this.resolveDestinationRoot();
    let dbPath = requestedAccessPath || this.session.accessPath || null;
    if (!dbPath) {
      dbPath = await this.detectAccessPath({ accessPath: requestedAccessPath });
    }
    if (!dbPath) throw new Error("No hay BD detectada para esta operación.");

    const modulesPath = this.modulesPathFor(dbPath, destinationRootAbs);

    const needsSessionUpdate =
      allowSessionMutation &&
      this.session.active &&
      (
        this.session.accessPath !== dbPath ||
        this.session.destinationRoot !== destinationRootAbs ||
        this.session.modulesPath !== modulesPath
      );

    if (needsSessionUpdate) {
      this.session.accessPath = dbPath;
      this.session.destinationRoot = destinationRootAbs;
      this.session.modulesPath = modulesPath;
      await this.saveSessionToDisk();
    }

    return {
      accessPath: dbPath,
      destinationRootAbs,
      modulesPath
    };
  }

  async recordSessionImport(mods) {
    if (!this.session.active) return;
    const now = new Date().toISOString();
    this.session.lastSyncAt = now;
    this.session.changedModules = uniq([...(this.session.changedModules || []), ...(mods || [])]);
    this.session.pendingModules = [];
    await this.saveSessionToDisk();
  }

  runVbaManager({ action, accessPath, destinationRootAbs, moduleNames = [], backendPath, erdPath, location, importMode, json = false, procedureName, procedureArgs = [] }) {
    return new Promise((resolve, reject) => {
      const exe = powershellExe();
      const args = [
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        this.vbaManagerPath,
        "-Action",
        action,
        "-DestinationRoot",
        destinationRootAbs
      ];

      if (accessPath) {
        args.push("-AccessPath", accessPath);
      }

      if (backendPath) {
        args.push("-BackendPath", backendPath);
      }
      if (erdPath) {
        args.push("-ErdPath", erdPath);
      }
      if (location) {
        args.push("-Location", location);
      }

      if (importMode) {
        args.push("-ImportMode", importMode);
      }

      if (procedureName) {
        args.push("-ProcedureName", procedureName);
      }

      if (Array.isArray(procedureArgs) && procedureArgs.length > 0) {
        args.push("-ProcedureArgsJson", JSON.stringify(procedureArgs));
      }

      if (json) {
        args.push("-Json");
      }

      if (this.allowStartupExecution) {
        args.push("-AllowStartupExecution");
      }

      // Pasar los nombres como JSON evita romper módulos válidos con comas
      // y elimina la ambigüedad posicional del binding de PowerShell.
      if (Array.isArray(moduleNames) && moduleNames.length > 0) {
        args.push("-ModuleNamesJson", JSON.stringify(moduleNames));
      }

      if (this.password) {
        if (this.passwordSource === "cli" && !this._passwordArgWarningShown) {
          console.warn("⚠️  --password se pasa como argumento de proceso y puede ser visible para el SO. Preferí ACCESS_VBA_PASSWORD.");
          this._passwordArgWarningShown = true;
        }
        args.push("-Password", this.password);
      }

      const child = spawn(exe, args, {
        cwd: this.projectRoot,
        windowsHide: true
      });

      let stdout = "";
      let stderr = "";
      child.stdout.on("data", (d) => {
        stdout += d.toString();
      });
      child.stderr.on("data", (d) => {
        stderr += d.toString();
      });

      child.on("error", reject);
      child.on("close", (code) => {
        if (code === 0) return resolve({ code, stdout, stderr });
        const err = new Error(`VBAManager.ps1 terminó con código ${code}`);
        err.code = code;
        err.stdout = stdout;
        err.stderr = stderr;
        return reject(err);
      });
    });
  }

  looksLikeModuleArrayNotSupported(err) {
    const stderr = String((err && err.stderr) || "");
    const stdout = String((err && err.stdout) || "");
    const t = (stderr + "\n" + stdout).toLowerCase();
    return (
      t.includes("cannot process argument transformation") ||
      t.includes("cannot convert") && (t.includes("modulenames") || t.includes("modulenamesjson")) ||
      t.includes("parameterbinding") && (t.includes("modulenames") || t.includes("modulenamesjson") || t.includes("argument transformation"))
    );
  }

  async backupSrcIfExists() {
    const srcPath = this.resolveDestinationRoot();
    let hasContent = false;
    try {
      const entries = await fsp.readdir(srcPath, { withFileTypes: true });
      hasContent = entries.some((e) => e.isFile() || e.isDirectory());
    } catch {
      return;
    }
    if (!hasContent) return;

    const backupPath = srcPath + ".bak";
    try {
      await fsp.access(backupPath);
      console.log(`⚠️  Ya existía ${path.basename(backupPath)}/; se sobrescribirá con un backup nuevo.`);
      await fsp.rm(backupPath, { recursive: true, force: true });
    } catch {}
    console.log(`💾 Backup de src/ → ${path.basename(backupPath)}/`);
    await copyDirRecursive(srcPath, backupPath);
  }

  async start({ accessPath } = {}) {
    await this.initializeSession({ accessPath, performInitialExport: true });
  }

  resolveDocumentArtifacts(moduleName) {
    const variants = logicalModuleNameVariants(moduleName);
    const sourceRoot = this.resolveDestinationRoot();
    const folders = [
      { root: path.join(sourceRoot, "forms"), textExt: ".form.txt" },
      { root: path.join(sourceRoot, "reports"), textExt: ".report.txt" }
    ];

    for (const folder of folders) {
      for (const variant of variants) {
        const textPath = path.join(folder.root, variant + folder.textExt);
        if (!fs.existsSync(textPath)) continue;

        const baseVariant = String(variant).replace(/^(Form|Report)_/i, "");
        const clsResolution = pickBestDocumentClsPath({
          canonicalPath: path.join(folder.root, baseVariant + ".cls"),
          legacyPath: path.join(folder.root, variant + ".cls")
        });

        return {
          moduleName,
          textPath,
          clsPath: clsResolution.clsPath,
          mirrorClsPaths: clsResolution.mirrorPaths,
          kind: folder.textExt === ".report.txt" ? "report" : "form",
          clsWarning: clsResolution.warning,
          clsError: clsResolution.error
        };
      }
    }

    return null;
  }

  async syncCodeBehind(moduleNames) {
    const mods = uniq((moduleNames || []).map(String).filter(Boolean));
    let synced = 0;
    
    for (const mod of mods) {
      const artifacts = this.resolveDocumentArtifacts(mod);
      if (!artifacts || !artifacts.clsPath || !artifacts.textPath) continue;
      if (artifacts.clsError) {
        throw new Error(artifacts.clsError);
      }
      if (artifacts.clsWarning) {
        console.log(`  ${artifacts.clsWarning}`);
      }
      
      let formContent;
      try {
        formContent = await fsp.readFile(artifacts.textPath, "utf8");
      } catch (err) {
        console.warn(`WARN: no se pudo leer ${artifacts.textPath}; se omite sync de CodeBehind para ${mod}: ${err && err.message ? err.message : String(err)}`);
        continue;
      }

      let clsContent;
      try {
        clsContent = await fsp.readFile(artifacts.clsPath, "utf8");
      } catch (err) {
        console.warn(`WARN: no se pudo leer ${artifacts.clsPath}; se omite sync de CodeBehind para ${mod}: ${err && err.message ? err.message : String(err)}`);
        continue;
      }

      for (const mirrorPath of artifacts.mirrorClsPaths || []) {
        try {
          await fsp.writeFile(mirrorPath, clsContent, "utf8");
          console.log(`  🔁 Unificado sidecar: ${path.basename(mirrorPath)} ← ${path.basename(artifacts.clsPath)}`);
        } catch (err) {
          console.warn(`WARN: no se pudo unificar sidecar ${mirrorPath}; puede quedar divergente respecto de ${artifacts.clsPath}: ${err && err.message ? err.message : String(err)}`);
        }
      }

      const newFormContent = mergeDocumentCodeBehindText(formContent, clsContent);

      await fsp.writeFile(artifacts.textPath, newFormContent, "utf8");
      console.log(`  🔄 Sincronizado CodeBehind: ${path.basename(artifacts.textPath)} ← ${path.basename(artifacts.clsPath)}`);
      synced++;
    }

    return synced;
  }

  async syncAllDocumentCodeBehind() {
    const sourceRoot = this.resolveDestinationRoot();
    const folders = [
      path.join(sourceRoot, "forms"),
      path.join(sourceRoot, "reports")
    ];
    const modules = [];

    for (const folder of folders) {
      if (!fs.existsSync(folder)) continue;
      for (const entry of fs.readdirSync(folder, { withFileTypes: true })) {
        if (!entry.isFile()) continue;
        const lower = entry.name.toLowerCase();
        if (!(lower.endsWith(".form.txt") || lower.endsWith(".report.txt"))) continue;
        modules.push(entry.name.replace(/(\.form|\.report)\.txt$/i, ""));
      }
    }

    return this.syncCodeBehind(modules);
  }

  async importModules({ moduleNames, accessPath, importMode = "Auto", trackSession = false } = {}) {
    const context = await this.resolveCommandContext({ accessPath, initializeIfNeeded: trackSession, allowSessionMutation: trackSession });
    const destinationRootAbs = context.destinationRootAbs;
    const dbPath = context.accessPath;

    const mods = uniq((moduleNames || []).map(String).filter(Boolean));
    if (mods.length === 0) throw new Error("No se especificaron módulos para importar.");

    // ✅ Sync CodeBehind antes de importar código (.cls)
    if (importMode === "Code" || importMode === "Auto") {
      await this.syncCodeBehind(mods);
    }

    console.log(`🔄 Importando ${mods.length} módulo(s): ${mods.join(", ")}`);

    let vbaResult;
    try {
      vbaResult = await this.runVbaManager({
        action: "Import",
        accessPath: dbPath,
        destinationRootAbs,
        moduleNames: mods,
        importMode
      });
    } catch (err) {
      logModuleResults(parseModuleResults(err.stdout));
      if (mods.length > 1 && this.looksLikeModuleArrayNotSupported(err)) {
        console.warn(`⚠️  Fallback: se reintenta módulo a módulo por incompatibilidad del binding de arrays con PowerShell. Error original: ${err && err.message ? err.message : String(err)}`);
        const failedModules = [];
        for (const m of mods) {
          try {
            const singleResult = await this.runVbaManager({
              action: "Import",
              accessPath: dbPath,
              destinationRootAbs,
              moduleNames: [m],
              importMode
            });
            logModuleResults(parseModuleResults(singleResult.stdout));
          } catch (singleErr) {
            failedModules.push({ module: m, error: singleErr });
            const parsed = parseModuleResults(singleErr.stdout);
            if (parsed) {
              logModuleResults(parsed);
            } else {
              console.warn(`❌ ${m}: falló import individual: ${singleErr && singleErr.message ? singleErr.message : String(singleErr)}`);
            }
          }
        }
        if (failedModules.length > 0) {
          const summary = failedModules.map((item) => `${item.module}: ${item.error && item.error.message ? item.error.message : String(item.error)}`).join("; ");
          const fallbackErr = new Error(`Falló el import individual de ${failedModules.length}/${mods.length} módulo(s): ${summary}`);
          fallbackErr.failures = failedModules;
          throw fallbackErr;
        }
      } else {
        throw err;
      }
    }
    if (vbaResult) logModuleResults(parseModuleResults(vbaResult.stdout));

    if (trackSession) {
      await this.recordSessionImport(mods);
    }

    console.log("✅ Sync terminado.");
    console.log("Abre Access → VBE → Debug → Compile");
  }

  async importForms({ moduleNames, accessPath } = {}) {
    return this.importModules({ moduleNames, accessPath, importMode: "Form" });
  }

  async importCode({ moduleNames, accessPath } = {}) {
    return this.importModules({ moduleNames, accessPath, importMode: "Code" });
  }

  async flushPendingImports() {
    if (this.importing) return;
    const mods = uniq([...this.pendingModules]);
    if (mods.length === 0) return;

    const now = Date.now();
    const ready = [];
    const blocked = [];
    let minRemainingMs = null;

    for (const m of mods) {
      const last = this.lastImportedAtByModule.get(m) || 0;
      const elapsed = now - last;
      if (elapsed < this.cooldownMs) {
        const remaining = this.cooldownMs - elapsed;
        blocked.push(m);
        if (minRemainingMs == null || remaining < minRemainingMs) minRemainingMs = remaining;
      } else {
        ready.push(m);
      }
    }

    this.pendingModules = new Set(blocked);

    if (ready.length === 0) {
      if (minRemainingMs != null) this.scheduleDebouncedImport(Math.max(50, minRemainingMs));
      return;
    }

    const dbPath = this.session.accessPath;
    if (dbPath) {
      const ext = path.extname(dbPath).toLowerCase();
      const lockExt = ext === ".mdb" || ext === ".mde" ? ".ldb" : ".laccdb";
      const lockPath = dbPath.slice(0, -ext.length) + lockExt;
      if (fs.existsSync(lockPath)) {
        for (const m of ready) this.pendingModules.add(m);
        if (!this._accessOpenDeferred) {
          console.log(`⏸️  Access abierto (${path.basename(dbPath)}), import postergado: ${ready.join(", ")}`);
          console.log("   Se importará automáticamente cuando cierres Access.");
          this._accessOpenDeferred = true;
        }
        this.scheduleDebouncedImport(5000);
        return;
      }
    }
    if (this._accessOpenDeferred) {
      console.log(`▶️  Access cerrado. Importando: ${ready.join(", ")}`);
      this._accessOpenDeferred = false;
    }

    this.importing = true;
    try {
      await this.importModules({ moduleNames: ready, trackSession: true });
      const doneAt = Date.now();
      for (const m of ready) this.lastImportedAtByModule.set(m, doneAt);
    } finally {
      this.importing = false;
      if (this.pendingModules.size > 0) this.scheduleDebouncedImport();
    }
  }

  scheduleDebouncedImport(delayMs) {
    if (this.debounceTimer) clearTimeout(this.debounceTimer);

    this.session.pendingModules = uniq([...(this.session.pendingModules || []), ...this.pendingModules]);
    this.saveSessionToDisk().catch((err) => {
      console.warn(`WARN: no se pudo guardar session.json (${this.stateFile}): ${err && err.message ? err.message : String(err)}`);
    });

    this.debounceTimer = setTimeout(async () => {
      this.debounceTimer = null;
      try {
        await this.flushPendingImports();
      } catch (err) {
        console.error("ERROR importando lote:");
        console.error(err && err.message ? err.message : String(err));
        if (err && err.stdout) console.error("\n--- stdout ---\n" + err.stdout);
        if (err && err.stderr) console.error("\n--- stderr ---\n" + err.stderr);
      }
    }, Number.isFinite(delayMs) ? delayMs : this.debounceMs);
  }

  async watch({ accessPath } = {}) {
    await this.ensureReady();
    await this.loadSessionFromDisk();

    if (!this.session.active) {
      await this.start({ accessPath });
      await this.loadSessionFromDisk();
    }

    const modulesPath = this.session.modulesPath;
    if (!modulesPath) throw new Error("No se pudo determinar modulesPath.");
    await fsp.mkdir(modulesPath, { recursive: true });

    this.session.watcherPid = process.pid;
    await this.saveSessionToDisk();

    console.log("👀 Watcher activo en:");
    console.log(`   ${modulesPath}\n`);
    console.log("   (Ctrl+C para terminar: se cerrará la sesión)\n");

    const chokidar = require("chokidar");
    this.watcher = chokidar.watch(modulesPath, {
      ignoreInitial: true,
      awaitWriteFinish: { stabilityThreshold: 800, pollInterval: 100 }
    });

    const onTouched = (filePath) => {
      if (!isWatchedExt(filePath)) return;
      try {
        const st = fs.statSync(filePath);
        const sig = `${st.size}:${st.mtimeMs}`;
        const prev = this.fileSignatures.get(filePath);
        if (prev === sig) return;
        this.fileSignatures.set(filePath, sig);
      } catch {
        return;
      }
      const mod = moduleNameFromFile(filePath);
      if (!mod) return;
      this.pendingModules.add(mod);
      this.scheduleDebouncedImport();
    };

    this.watcher.on("add", onTouched);
    this.watcher.on("change", onTouched);
    this.watcher.on("unlink", (filePath) => {
      if (!isWatchedExt(filePath)) return;
      console.log(`⚠️  Archivo eliminado: ${path.basename(filePath)} (no se borra el módulo en Access automáticamente)`);
    });

    const shutdown = async () => {
      try {
        await this.end();
      } finally {
        process.exit(0);
      }
    };

    process.on("SIGINT", shutdown);
    process.on("SIGTERM", shutdown);
  }

  async stopWatching() {
    if (this.debounceTimer) {
      clearTimeout(this.debounceTimer);
      this.debounceTimer = null;
    }
    if (this.watcher) {
      await this.watcher.close();
      this.watcher = null;
    }
  }

  async stopWatcherProcessIfDifferentPid() {
    await this.loadSessionFromDisk();
    const rawPid = this.session.watcherPid;
    if (!rawPid) return;
    const pid = Number(rawPid);
    if (!Number.isInteger(pid) || pid <= 0) {
      console.log(`⚠️  watcherPid inválido en session.json; se ignora por seguridad.`);
      this.session.watcherPid = null;
      await this.saveSessionToDisk();
      return;
    }
    if (pid === process.pid) return;

    try {
      process.kill(pid, 0);
    } catch {
      this.session.watcherPid = null;
      await this.saveSessionToDisk();
      return;
    }

    try {
      if (process.platform === "win32") {
        require("child_process").execSync(`taskkill /PID ${pid} /F /T`, { stdio: "ignore" });
      } else {
        process.kill(pid, "SIGTERM");
      }
      console.log(`🛑 Watcher detenido (pid ${pid})`);
    } catch (err) {
      console.log(`⚠️  No se pudo detener el watcher (pid ${pid}): ${err && err.message ? err.message : String(err)}`);
    }
  }

  async end() {
    await this.ensureReady();
    await this.loadSessionFromDisk();

    if (!this.session.active) {
      console.log("ℹ️  No hay sesión activa.");
      return;
    }

    await this.stopWatcherProcessIfDifferentPid();
    await this.stopWatching();

    const pending = uniq([...(this.session.pendingModules || []), ...this.pendingModules]);
    this.pendingModules.clear();
    this.session.pendingModules = [];
    await this.saveSessionToDisk();

    if (pending.length > 0) {
      console.log(`🔄 Sync final de pendientes: ${pending.join(", ")}`);
      await this.importModules({ moduleNames: pending, trackSession: true });
    }

    if (this.autoExportOnEnd) {
      console.log("📦 Export final...");
      await this.runVbaManager({
        action: "Export",
        accessPath: this.session.accessPath,
        destinationRootAbs: this.session.destinationRoot
      });
      console.log("✓ Export final completado");
    } else {
      console.log("ℹ️  Export final desactivado (auto_export_on_end=false)");
    }

    const changed = uniq(this.session.changedModules || []);
    this.session.active = false;
    this.session.watcherPid = null;
    await this.saveSessionToDisk();

    console.log("\n🏁 Sesión finalizada.");
    console.log(`   Módulos sincronizados: ${changed.length}`);
    if (changed.length > 0) console.log("   " + changed.join(", "));
  }

  async fixEncoding({ moduleNames, accessPath, location = "Both" } = {}) {
    const context = await this.resolveCommandContext({ accessPath, initializeIfNeeded: false, allowSessionMutation: false });
    const destinationRootAbs = context.destinationRootAbs;
    const dbPath = context.accessPath;

    const mods = moduleNames && moduleNames.length > 0
      ? uniq(moduleNames.map(String).filter(Boolean))
      : [];

    const desc = mods.length > 0 ? mods.join(", ") : "todos los módulos";
    console.log(`🔧 Fix-encoding (location: ${location}) — ${desc}...`);

    await this.runVbaManager({
      action: "Fix-Encoding",
      accessPath: dbPath,
      destinationRootAbs,
      moduleNames: mods,
      location
    });

    if (this.session.active) {
      await this.recordSessionImport(mods.length > 0 ? mods : ["fix-encoding:*"]);
    }

    console.log("✓ Fix-encoding completado.");
  }

  async exportModules({ moduleNames, accessPath } = {}) {
    const context = await this.resolveCommandContext({ accessPath, initializeIfNeeded: false, allowSessionMutation: false });
    const destinationRootAbs = context.destinationRootAbs;
    const dbPath = context.accessPath;

    const mods = uniq((moduleNames || []).map(String).filter(Boolean));
    if (mods.length === 0) throw new Error("No se especificaron módulos para exportar.");

    console.log(`📤 Exportando ${mods.length} módulo(s): ${mods.join(", ")}`);

    await this.runVbaManager({
      action: "Export",
      accessPath: dbPath,
      destinationRootAbs,
      moduleNames: mods
    });

    console.log("✓ Export completado.");
  }

  async exportAll({ accessPath } = {}) {
    const context = await this.resolveCommandContext({ accessPath, initializeIfNeeded: false, allowSessionMutation: false });
    const detectedAccess = context.accessPath;
    const destinationRootAbs = context.destinationRootAbs;

    await this.backupSrcIfExists();
    console.log("📤 Exportando todos los módulos...");
    await this.runVbaManager({
      action: "Export",
      accessPath: detectedAccess,
      destinationRootAbs
    });
    console.log("✓ Export completado.");
  }

  async importAll({ accessPath } = {}) {
    const context = await this.resolveCommandContext({ accessPath, initializeIfNeeded: false, allowSessionMutation: false });
    const destinationRootAbs = context.destinationRootAbs;
    const dbPath = context.accessPath;

    await this.syncAllDocumentCodeBehind();

    const backupPath = dbPath + ".bak";
    try {
      await fsp.access(backupPath);
      console.log(`⚠️  Ya existía ${path.basename(backupPath)}; se sobrescribirá con un backup nuevo.`);
      await fsp.rm(backupPath, { force: true });
    } catch {}
    console.log(`💾 Backup: ${path.basename(backupPath)}`);
    await fsp.copyFile(dbPath, backupPath);

    console.log("📥 Importando todos los módulos desde src/...");
    try {
      await this.runVbaManager({
        action: "Import",
        accessPath: dbPath,
        destinationRootAbs
      });
    } catch (err) {
      console.error("❌ Import fallido. Restaurando backup...");
      try {
        await fsp.rm(dbPath, { force: true });
        await fsp.rename(backupPath, dbPath);
        console.log("✅ Base de datos restaurada al estado previo.");
      } catch (restoreErr) {
        console.error(`⚠️  ERROR CRÍTICO: no se pudo restaurar el backup (${backupPath}). Restaura manualmente.`);
        console.error(restoreErr && restoreErr.message ? restoreErr.message : String(restoreErr));
      }
      throw err;
    }

    try {
      await fsp.unlink(backupPath);
      console.log("✅ Import-all completado. Backup eliminado.");
    } catch (err) {
      console.log(`⚠️  Import-all completado, pero no se pudo eliminar el backup ${path.basename(backupPath)}: ${err && err.message ? err.message : String(err)}`);
    }
    console.log("Abre Access → VBE → Debug → Compile");
  }

  async generateErd({ backendPath, erdPath } = {}) {
    await this.ensureReady();
    const destinationRootAbs = this.resolveDestinationRoot();

    console.log("📊 Generando ERD...");
    await this.runVbaManager({
      action: "Generate-ERD",
      destinationRootAbs,
      backendPath,
      erdPath
    });
    console.log("✅ ERD Generado.");
  }

  async listObjects({ accessPath, json = false } = {}) {
    const context = await this.resolveCommandContext({ accessPath, initializeIfNeeded: false, allowSessionMutation: false });
    const destinationRootAbs = context.destinationRootAbs;
    const dbPath = context.accessPath;

    const result = await this.runVbaManager({
      action: "List-Objects",
      accessPath: dbPath,
      destinationRootAbs,
      json
    });

    if (json) {
      return parseJsonFromStdout(result.stdout, "JSON de List-Objects");
    }

    process.stdout.write(result.stdout || "");
    return null;
  }

  async exists({ moduleName, accessPath, json = false } = {}) {
    const context = await this.resolveCommandContext({ accessPath, initializeIfNeeded: false, allowSessionMutation: false });
    const destinationRootAbs = context.destinationRootAbs;
    const dbPath = context.accessPath;
    if (!moduleName) throw new Error("Falta moduleName para exists.");

    const result = await this.runVbaManager({
      action: "Exists",
      accessPath: dbPath,
      destinationRootAbs,
      moduleNames: [moduleName],
      json
    });

    if (json) {
      return parseJsonFromStdout(result.stdout, "JSON de Exists");
    }

    process.stdout.write(result.stdout || "");
    return null;
  }

  async runVba({ procedureName, argsJson, accessPath, json = false } = {}) {
    const context = await this.resolveCommandContext({ accessPath, initializeIfNeeded: false, allowSessionMutation: false });
    const destinationRootAbs = context.destinationRootAbs;
    const dbPath = context.accessPath;
    const name = String(procedureName || "").trim();
    if (!name) throw new Error("Falta procedureName para run-vba.");
    const procedureArgs = parseVbaArgsJson(argsJson);

    const result = await this.runVbaManager({
      action: "Run-Procedure",
      accessPath: dbPath,
      destinationRootAbs,
      procedureName: name,
      procedureArgs,
      json: true,
    });

    const parsed = parseJsonFromStdout(result.stdout, "JSON de Run-Procedure");
    if (json) {
      return parsed;
    }

    if (parsed.ok) {
      const rendered = parsed.returnValue === null || parsed.returnValue === undefined ? "" : String(parsed.returnValue);
      console.log(`✅ ${parsed.procedure} ejecutado correctamente${rendered ? `: ${rendered}` : ""}`);
    } else {
      console.error(`❌ ${parsed.procedure} falló: ${parsed.error || "error desconocido"}`);
    }
    if (Array.isArray(parsed.logs) && parsed.logs.length > 0) {
      console.log("📋 Logs:");
      for (const line of parsed.logs) console.log(`  ${line}`);
    }
    return parsed;
  }

  async compileVba({ accessPath, json = false } = {}) {
    const context = await this.resolveCommandContext({ accessPath, initializeIfNeeded: false, allowSessionMutation: false });
    const destinationRootAbs = context.destinationRootAbs;
    const dbPath = context.accessPath;

    const result = await this.runVbaManager({
      action: "Compile",
      accessPath: dbPath,
      destinationRootAbs,
      json: true
    });

    const parsed = parseJsonFromStdout(result.stdout, "JSON de Compile");
    if (json) return parsed;

    if (parsed.ok) {
      console.log("✅ Compilación VBA correcta.");
    } else {
      console.error(`❌ Compilación VBA fallida: ${parsed.error || "error desconocido"}`);
      if (parsed.component) console.error(`   Componente: ${parsed.component}`);
      if (parsed.line) console.error(`   Línea: ${parsed.line}${parsed.column ? `, columna: ${parsed.column}` : ""}`);
      if (parsed.sourceLine) console.error(`   Código: ${parsed.sourceLine}`);
    }
    return parsed;
  }

  async loadVbaTestPlan(testsPath) {
    const resolved = path.resolve(this.projectRoot, testsPath || "tests.vba.json");
    const content = await fsp.readFile(resolved, "utf8");
    let parsed;
    try {
      parsed = JSON.parse(content);
    } catch (err) {
      throw new Error(`No se pudo parsear ${resolved}: ${err && err.message ? err.message : String(err)}`);
    }
    const tests = Array.isArray(parsed) ? parsed : parsed.tests;
    if (!Array.isArray(tests)) {
      throw new Error(`${resolved} debe contener un array o un objeto con propiedad "tests".`);
    }
    return {
      path: resolved,
      tests: tests.map((test, index) => {
        if (!test || typeof test !== "object") throw new Error(`Test #${index + 1} inválido: debe ser un objeto.`);
        const procedure = String(test.procedure || test.proc || "").trim();
        if (!procedure) throw new Error(`Test #${index + 1} inválido: falta "procedure".`);
        const args = Object.prototype.hasOwnProperty.call(test, "args") ? test.args : [];
        if (!Array.isArray(args)) throw new Error(`Test #${index + 1} (${procedure}) inválido: "args" debe ser array.`);
        return {
          name: String(test.name || procedure),
          procedure,
          args,
          expect: test.expect || {},
          tags: Array.isArray(test.tags) ? test.tags.map(String) : []
        };
      })
    };
  }

  buildSingleProcedureTestPlan({ procedureName, argsJson } = {}) {
    const procedure = String(procedureName || "").trim();
    if (!procedure) throw new Error("Falta procedureName para test-vba directo.");
    return {
      path: null,
      tests: [{
        name: procedure,
        procedure,
        args: parseVbaArgsJson(argsJson),
        expect: {},
        tags: ["direct"]
      }]
    };
  }

  async testVba({ testsPath, procedureName, argsJson, filter, accessPath, json = false, compile = true } = {}) {
    const plan = procedureName
      ? this.buildSingleProcedureTestPlan({ procedureName, argsJson })
      : await this.loadVbaTestPlan(testsPath);
    const filterText = filter ? String(filter).toLowerCase() : null;
    const selected = filterText
      ? plan.tests.filter((test) =>
          test.name.toLowerCase().includes(filterText) ||
          test.procedure.toLowerCase().includes(filterText) ||
          test.tags.some((tag) => tag.toLowerCase().includes(filterText))
        )
      : plan.tests;

    if (compile) {
      const compileResult = await this.compileVba({ accessPath, json: true });
      if (!compileResult.ok) {
        const report = {
          ok: false,
          phase: "compile",
          testsPath: plan.path,
          total: selected.length,
          passed: 0,
          failed: 0,
          skipped: selected.length,
          compile: compileResult,
          results: []
        };
        if (!json) {
          console.error(`❌ No se ejecutan tests: falló compilación VBA.`);
          if (compileResult.component) console.error(`   Componente: ${compileResult.component}`);
          if (compileResult.line) console.error(`   Línea: ${compileResult.line}`);
          if (compileResult.sourceLine) console.error(`   Código: ${compileResult.sourceLine}`);
        }
        return report;
      }
    }

    const results = [];
    for (const test of selected) {
      let runResult;
      try {
        runResult = await this.runVba({
          procedureName: test.procedure,
          argsJson: JSON.stringify(test.args),
          accessPath,
          json: true
        });
      } catch (err) {
        runResult = {
          ok: false,
          procedure: test.procedure,
          argsCount: test.args.length,
          returnValue: null,
          returnType: null,
          payload: null,
          logs: [],
          error: err && err.message ? err.message : String(err)
        };
      }
      const assertion = evaluateVbaTestExpectation(runResult, test.expect);
      results.push({
        name: test.name,
        procedure: test.procedure,
        args: test.args,
        ok: assertion.ok,
        failures: assertion.failures,
        run: runResult,
        logs: Array.isArray(runResult.logs) ? runResult.logs : []
      });
      if (!json) {
        const icon = assertion.ok ? "✅" : "❌";
        console.log(`${icon} ${test.name} (${test.procedure})`);
        for (const line of runResult.logs || []) console.log(`   ${line}`);
        for (const failure of assertion.failures) console.log(`   ${failure}`);
      }
    }

    const failed = results.filter((result) => !result.ok).length;
    return {
      ok: failed === 0,
      phase: "tests",
      testsPath: plan.path,
      total: selected.length,
      passed: selected.length - failed,
      failed,
      skipped: plan.tests.length - selected.length,
      results
    };
  }

  async status() {
    await this.ensureReady();
    await this.loadSessionFromDisk();

    if (!this.session.active) {
      console.log("ℹ️  Estado: SIN SESIÓN");
      return;
    }

    const changed = uniq(this.session.changedModules || []);
    console.log("📌 Estado de sesión:");
    console.log(`   Activa: ${this.session.active}`);
    console.log(`   Iniciada: ${this.session.startedAt}`);
    console.log(`   Access: ${this.session.accessPath}`);
    console.log(`   DestinationRoot: ${this.session.destinationRoot}`);
    console.log(`   ModulesPath: ${this.session.modulesPath}`);
    console.log(`   Último sync: ${this.session.lastSyncAt || "—"}`);
    console.log(`   Pendientes: ${uniq(this.session.pendingModules || []).length}`);
    console.log(`   Módulos tocados: ${changed.length}`);
  }

  async _allDocumentModuleNames() {
    const sourceRoot = this.resolveDestinationRoot();
    const folders = [
      { dir: path.join(sourceRoot, "forms"), ext: ".form.txt" },
      { dir: path.join(sourceRoot, "reports"), ext: ".report.txt" }
    ];
    const names = [];
    for (const { dir, ext } of folders) {
      if (!fs.existsSync(dir)) continue;
      for (const entry of await fsp.readdir(dir, { withFileTypes: true })) {
        if (!entry.isFile() || !entry.name.toLowerCase().endsWith(ext)) continue;
        names.push(entry.name.slice(0, -ext.length));
      }
    }
    return names;
  }

  async verifyCode({ moduleNames } = {}) {
    const mods = uniq((moduleNames || []).map(String).filter(Boolean));
    const targets = mods.length > 0 ? mods : await this._allDocumentModuleNames();

    if (targets.length === 0) {
      console.log("ℹ️  No se encontraron formularios/reportes con .cls asociado.");
      return true;
    }

    let allInSync = true;

    for (const mod of targets) {
      const artifacts = this.resolveDocumentArtifacts(mod);
      if (!artifacts) {
        console.log(`⚠️  ${mod}: no se encontró .form.txt ni .report.txt`);
        continue;
      }
      if (!artifacts.clsPath) {
        console.log(`ℹ️  ${mod}: sin .cls (solo UI, sin code-behind verificable)`);
        continue;
      }

      let formContent, clsContent;
      try { formContent = fs.readFileSync(artifacts.textPath, "utf8"); } catch (e) {
        console.log(`⚠️  ${mod}: no se pudo leer ${path.basename(artifacts.textPath)}: ${e.message}`);
        continue;
      }
      try { clsContent = fs.readFileSync(artifacts.clsPath, "utf8"); } catch (e) {
        console.log(`⚠️  ${mod}: no se pudo leer ${path.basename(artifacts.clsPath)}: ${e.message}`);
        continue;
      }

      const section = splitCodeBehindSection(normalizeNewlines(formContent, "\n"));
      if (!section) {
        console.log(`⚠️  ${mod}: el .form.txt no tiene sección CodeBehind`);
        continue;
      }

      const formBodyNorm = normalizeNewlines(splitVbaMetadataHeaderText(section.body).body, "\n").trimEnd();
      const clsBodyNorm = normalizeNewlines(splitVbaMetadataHeaderText(sanitizeVbaImportText(clsContent)).body, "\n").trimEnd();

      if (formBodyNorm === clsBodyNorm) {
        console.log(`✅ ${mod}: en sync`);
      } else {
        allInSync = false;
        console.log(`❌ ${mod}: DESINCRONIZADO`);
        const diff = unifiedDiff(
          formBodyNorm.split("\n"),
          clsBodyNorm.split("\n"),
          path.basename(artifacts.textPath) + " (CodeBehind)",
          path.basename(artifacts.clsPath)
        );
        if (diff) console.log(diff);
      }
    }

    return allInSync;
  }
}

module.exports = {
  AccessVbaSyncSkill,
  _test: {
    mergeDocumentCodeBehindText,
    splitCodeBehindSection,
    sanitizeVbaImportText,
    normalizeNewlines,
    normalizePathForComparison,
    splitVbaMetadataHeaderText,
    moduleNameFromFile,
    isWatchedExt,
    logicalModuleNameVariants,
    pickBestDocumentClsPath,
    unifiedDiff,
    parseModuleResults,
    parseJsonFromStdout,
    parseVbaArgsJson,
    evaluateVbaTestExpectation,
    countMeaningfulVbaBodyLines,
  }
};
