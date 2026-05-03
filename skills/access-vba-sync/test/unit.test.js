"use strict";

const { test, describe } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const {
  _test: {
    mergeDocumentCodeBehindText,
    splitCodeBehindSection,
    splitVbaMetadataHeaderText,
    sanitizeVbaImportText,
    normalizeNewlines,
    normalizePathForComparison,
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
  },
} = require("../handler");
const { _test: { parseArgs } } = require("../cli");

// ---------------------------------------------------------------------------
// moduleNameFromFile
// ---------------------------------------------------------------------------
describe("moduleNameFromFile", () => {
  test(".form.txt quita sufijo completo", () => {
    assert.equal(moduleNameFromFile("src/forms/Form_frmGestion.form.txt"), "Form_frmGestion");
  });
  test(".bas quita extension", () => {
    assert.equal(moduleNameFromFile("src/modules/Utilidades.bas"), "Utilidades");
  });
  test(".cls quita extension", () => {
    assert.equal(moduleNameFromFile("src/classes/CUsuario.cls"), "CUsuario");
  });
  test("nombre sin directorio", () => {
    assert.equal(moduleNameFromFile("ModuloX.bas"), "ModuloX");
  });
  test(".form.txt en Windows con backslash", () => {
    assert.equal(moduleNameFromFile("src\\forms\\Form_A.form.txt"), "Form_A");
  });
});

// ---------------------------------------------------------------------------
// isWatchedExt
// ---------------------------------------------------------------------------
describe("isWatchedExt", () => {
  test("detecta .form.txt", () => assert.ok(isWatchedExt("Form_A.form.txt")));
  test("detecta .bas", () => assert.ok(isWatchedExt("Mod.bas")));
  test("detecta .cls", () => assert.ok(isWatchedExt("Class.cls")));
  test("detecta .frm", () => assert.ok(isWatchedExt("Form.frm")));
  test("ignora .txt genérico", () => assert.ok(!isWatchedExt("notes.txt")));
  test("ignora .accdb", () => assert.ok(!isWatchedExt("BD.accdb")));
  test("case-insensitive en extensión", () => assert.ok(isWatchedExt("MOD.BAS")));
  test("case-insensitive en .FORM.TXT", () => assert.ok(isWatchedExt("Form_A.FORM.TXT")));
});

// ---------------------------------------------------------------------------
// logicalModuleNameVariants
// ---------------------------------------------------------------------------
describe("logicalModuleNameVariants", () => {
  test("nombre base genera variantes Form_ y Report_", () => {
    const v = logicalModuleNameVariants("frmGestion");
    assert.ok(v.includes("frmGestion"));
    assert.ok(v.includes("Form_frmGestion"));
    assert.ok(v.includes("Report_frmGestion"));
  });
  test("nombre con Form_ incluye el base sin prefijo", () => {
    const v = logicalModuleNameVariants("Form_frmGestion");
    assert.ok(v.includes("Form_frmGestion"));
    assert.ok(v.includes("frmGestion"));
  });
  test("sin duplicados en la salida", () => {
    const v = logicalModuleNameVariants("Form_frmGestion");
    assert.equal(v.length, new Set(v).size);
  });
  test("nombre vacío no incluye variantes con base vacía en el array", () => {
    const v = logicalModuleNameVariants("");
    assert.ok(!v.includes(""), "no debe incluir string vacío como variante");
  });
});

// ---------------------------------------------------------------------------
// splitCodeBehindSection
// ---------------------------------------------------------------------------
describe("splitCodeBehindSection", () => {
  test("encuentra marcador CodeBehind", () => {
    const text = "Begin Form\n  Width=100\nEnd\nCodeBehind\nOption Explicit\n";
    const s = splitCodeBehindSection(text);
    assert.ok(s !== null);
    assert.ok(s.before.includes("Begin Form"));
    assert.ok(s.body.includes("Option Explicit"));
    assert.equal(s.markerLine.trim(), "CodeBehind");
  });
  test("devuelve null cuando no hay marcador", () => {
    assert.equal(splitCodeBehindSection("Begin Form\n  Width=100\nEnd\n"), null);
  });
  test("acepta variante CodeBehindSection", () => {
    const text = "Begin Form\nEnd\nCodeBehindSection\nSub Foo()\nEnd Sub\n";
    const s = splitCodeBehindSection(text);
    assert.ok(s !== null);
    assert.ok(s.body.includes("Sub Foo()"));
  });
  test("before no incluye el marcador en sí", () => {
    const text = "UI\nCodeBehind\nCode\n";
    const s = splitCodeBehindSection(text);
    assert.ok(!s.before.includes("CodeBehind"));
  });
});

// ---------------------------------------------------------------------------
// sanitizeVbaImportText
// ---------------------------------------------------------------------------
describe("sanitizeVbaImportText", () => {
  test("elimina bloque VERSION CLASS / BEGIN / END", () => {
    const text = "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1\nEND\nOption Explicit\nSub Foo()\nEnd Sub\n";
    const out = sanitizeVbaImportText(text);
    assert.ok(!out.includes("VERSION"));
    assert.ok(!out.includes("BEGIN"));
    assert.ok(out.includes("Sub Foo()"));
  });
  test("elimina duplicados de Option Explicit", () => {
    const text = "Option Explicit\nOption Explicit\nSub Foo()\nEnd Sub\n";
    const out = sanitizeVbaImportText(text);
    const count = out.split("\n").filter((l) => l.trim() === "Option Explicit").length;
    assert.equal(count, 1);
  });
  test("elimina BOM del inicio", () => {
    const text = "﻿Option Explicit\nSub Foo()\nEnd Sub\n";
    const out = sanitizeVbaImportText(text);
    assert.ok(!out.startsWith("﻿"));
    assert.ok(out.includes("Option Explicit"));
  });
  test("elimina líneas Attribute VB_", () => {
    const text = "Attribute VB_Name = \"MyClass\"\nOption Explicit\nSub Foo()\nEnd Sub\n";
    const out = sanitizeVbaImportText(text);
    assert.ok(!out.includes("Attribute VB_Name"));
    assert.ok(out.includes("Option Explicit"));
  });
  test("preserva cuerpo del módulo intacto", () => {
    const body = "Sub Calcular(x As Long)\n  Dim y As Long\n  y = x * 2\nEnd Sub";
    const text = "VERSION 1.0 CLASS\nBEGIN\nEND\nOption Explicit\n" + body;
    const out = sanitizeVbaImportText(text);
    assert.ok(out.includes("Sub Calcular(x As Long)"));
    assert.ok(out.includes("y = x * 2"));
  });
});

// ---------------------------------------------------------------------------
// mergeDocumentCodeBehindText
// ---------------------------------------------------------------------------
describe("mergeDocumentCodeBehindText", () => {
  const formBase =
    "Version =21\nBegin Form\n  Width = 1000\nEnd\nCodeBehind\nOption Explicit\nSub OldSub()\nEnd Sub\n";
  const clsNew = "Option Explicit\nSub NewSub()\nEnd Sub\n";

  test("reemplaza cuerpo CodeBehind con contenido del .cls", () => {
    const result = mergeDocumentCodeBehindText(formBase, clsNew);
    assert.ok(result.includes("NewSub"));
    assert.ok(!result.includes("OldSub"));
  });
  test("preserva sección UI antes del marcador", () => {
    const result = mergeDocumentCodeBehindText(formBase, clsNew);
    const cbIdx = result.indexOf("CodeBehind");
    const ui = result.slice(0, cbIdx);
    assert.ok(ui.includes("Begin Form"));
    assert.ok(ui.includes("Width = 1000"));
  });
  test("lanza error cuando el form no tiene marcador CodeBehind", () => {
    const noMarker = "Begin Form\n  Width = 1000\nEnd\n";
    assert.throws(() => mergeDocumentCodeBehindText(noMarker, clsNew), /CodeBehind/);
  });
  test("el resultado contiene el marcador CodeBehind", () => {
    const result = mergeDocumentCodeBehindText(formBase, clsNew);
    assert.ok(result.includes("CodeBehind"));
  });
  test("el .cls vacío produce CodeBehind vacío pero form UI intacto", () => {
    const result = mergeDocumentCodeBehindText(formBase, "");
    assert.ok(result.includes("Begin Form"));
    assert.ok(!result.includes("OldSub"));
  });
});

// ---------------------------------------------------------------------------
// countMeaningfulVbaBodyLines
// ---------------------------------------------------------------------------
describe("countMeaningfulVbaBodyLines", () => {
  test("cuenta líneas de código reales", () => {
    const text = "Option Explicit\nSub Foo()\n  x = 1\nEnd Sub\n";
    assert.ok(countMeaningfulVbaBodyLines(text) > 0);
  });
  test("devuelve 0 para módulo vacío", () => {
    assert.equal(countMeaningfulVbaBodyLines(""), 0);
  });
  test("ignora comentarios", () => {
    const conComentarios = "Option Explicit\n' solo un comentario\n";
    const sinComentarios = "Option Explicit\n";
    assert.equal(
      countMeaningfulVbaBodyLines(conComentarios),
      countMeaningfulVbaBodyLines(sinComentarios)
    );
  });
});

// ---------------------------------------------------------------------------
// pickBestDocumentClsPath
// ---------------------------------------------------------------------------
describe("pickBestDocumentClsPath", () => {
  test("prioriza el sidecar canónico sin prefijo cuando el legacy está vacío", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "access-vba-sync-cls-"));
    const canonicalPath = path.join(dir, "subfrmDemo.cls");
    const legacyPath = path.join(dir, "Form_subfrmDemo.cls");

    fs.writeFileSync(canonicalPath, "Option Explicit\nSub Canonical()\nEnd Sub\n", "utf8");
    fs.writeFileSync(legacyPath, "Option Explicit\n", "utf8");

    const result = pickBestDocumentClsPath({ canonicalPath, legacyPath });
    assert.equal(result.clsPath, canonicalPath);
    assert.deepEqual(result.mirrorPaths, [legacyPath]);
    assert.ok(result.warning && result.warning.includes("vacío"));
  });

  test("si ambos divergen, prioriza el más recientemente modificado", async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "access-vba-sync-cls-"));
    const canonicalPath = path.join(dir, "subfrmDemo.cls");
    const legacyPath = path.join(dir, "Form_subfrmDemo.cls");

    fs.writeFileSync(canonicalPath, "Option Explicit\nSub CanonicalOld()\nEnd Sub\n", "utf8");
    fs.writeFileSync(legacyPath, "Option Explicit\nSub LegacyNew()\nEnd Sub\n", "utf8");
    fs.utimesSync(canonicalPath, new Date(1000), new Date(1000));
    fs.utimesSync(legacyPath, new Date(2000), new Date(2000));

    const result = pickBestDocumentClsPath({ canonicalPath, legacyPath });
    assert.equal(result.clsPath, legacyPath);
    assert.deepEqual(result.mirrorPaths, [canonicalPath]);
    assert.ok(result.warning && result.warning.includes("más recientemente"));
  });

  test("si ambos tienen el mismo cuerpo, usa el canónico y no alerta", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "access-vba-sync-cls-"));
    const canonicalPath = path.join(dir, "subfrmDemo.cls");
    const legacyPath = path.join(dir, "Form_subfrmDemo.cls");
    const content = "Option Explicit\nSub SameBody()\nEnd Sub\n";

    fs.writeFileSync(canonicalPath, content, "utf8");
    fs.writeFileSync(legacyPath, content, "utf8");

    const result = pickBestDocumentClsPath({ canonicalPath, legacyPath });
    assert.equal(result.clsPath, canonicalPath);
    assert.deepEqual(result.mirrorPaths, [legacyPath]);
    assert.equal(result.warning, null);
    assert.equal(result.error, null);
  });
});

// ---------------------------------------------------------------------------
// normalizePathForComparison
// ---------------------------------------------------------------------------
describe("normalizePathForComparison", () => {
  test("normaliza barras finales de forma estable", () => {
    const a = normalizePathForComparison("C:/Repo/Proyecto/");
    const b = normalizePathForComparison("C:/Repo/Proyecto");
    assert.equal(a, b);
  });
});

// ---------------------------------------------------------------------------
// unifiedDiff
// ---------------------------------------------------------------------------
describe("unifiedDiff", () => {
  test("líneas idénticas → null", () => {
    const lines = ["Línea A", "Línea B", "Línea C"];
    assert.equal(unifiedDiff(lines, lines, "a", "b"), null);
  });
  test("línea añadida aparece con +", () => {
    const a = ["Línea A", "Línea C"];
    const b = ["Línea A", "Línea B", "Línea C"];
    const diff = unifiedDiff(a, b, "a", "b");
    assert.ok(diff.includes("+Línea B"));
  });
  test("línea eliminada aparece con -", () => {
    const a = ["Línea A", "Línea B", "Línea C"];
    const b = ["Línea A", "Línea C"];
    const diff = unifiedDiff(a, b, "a", "b");
    assert.ok(diff.includes("-Línea B"));
  });
  test("incluye cabecera de hunk @@", () => {
    const diff = unifiedDiff(["A", "B"], ["A", "X"], "a", "b");
    assert.ok(diff.includes("@@"));
  });
  test("las etiquetas aparecen en el output", () => {
    const diff = unifiedDiff(["a"], ["b"], "file.form.txt", "file.cls");
    assert.ok(diff.includes("file.form.txt"));
    assert.ok(diff.includes("file.cls"));
  });
  test("líneas de contexto rodean los cambios", () => {
    const a = ["ctx1", "ctx2", "old", "ctx3", "ctx4"];
    const b = ["ctx1", "ctx2", "new", "ctx3", "ctx4"];
    const diff = unifiedDiff(a, b, "a", "b");
    assert.ok(diff.includes(" ctx2") || diff.includes(" ctx3"));
  });
  test("arrays vacíos → null", () => {
    assert.equal(unifiedDiff([], [], "a", "b"), null);
  });
  test("respeta contextLines y omite contexto lejano", () => {
    const a = ["pre1", "pre2", "old", "post1", "post2"];
    const b = ["pre1", "pre2", "new", "post1", "post2"];
    const diff = unifiedDiff(a, b, "a", "b", 1);
    assert.ok(!diff.includes(" pre1"));
    assert.ok(diff.includes(" pre2"));
    assert.ok(diff.includes("-old"));
    assert.ok(diff.includes("+new"));
    assert.ok(diff.includes(" post1"));
    assert.ok(!diff.includes(" post2"));
    assert.ok(diff.includes(" ..."));
  });
});

// ---------------------------------------------------------------------------
// parseModuleResults
// ---------------------------------------------------------------------------
describe("parseModuleResults", () => {
  test("parsea línea ##MODULE_RESULTS válida", () => {
    const stdout = 'Output\n##MODULE_RESULTS:[{"module":"Mod","status":"ok"}]\nMás output';
    const r = parseModuleResults(stdout);
    assert.ok(Array.isArray(r));
    assert.equal(r[0].module, "Mod");
    assert.equal(r[0].status, "ok");
  });
  test("devuelve null cuando no hay marcador", () => {
    assert.equal(parseModuleResults("Sin resultados aquí"), null);
  });
  test("devuelve null con JSON malformado", () => {
    assert.equal(parseModuleResults("##MODULE_RESULTS:no-es-json"), null);
  });
  test("parsea entradas con error", () => {
    const stdout = '##MODULE_RESULTS:[{"module":"A","status":"error","error":"falló"}]';
    const r = parseModuleResults(stdout);
    assert.equal(r[0].status, "error");
    assert.equal(r[0].error, "falló");
  });
  test("maneja string vacío y null", () => {
    assert.equal(parseModuleResults(""), null);
    assert.equal(parseModuleResults(null), null);
  });
  test("múltiples módulos en el array", () => {
    const stdout =
      '##MODULE_RESULTS:[{"module":"A","status":"ok"},{"module":"B","status":"error","error":"x"}]';
    const r = parseModuleResults(stdout);
    assert.equal(r.length, 2);
    assert.equal(r[1].module, "B");
  });
});

// ---------------------------------------------------------------------------
// normalizeNewlines
// ---------------------------------------------------------------------------
describe("normalizeNewlines", () => {
  test("convierte \\r\\n a \\n", () => {
    assert.equal(normalizeNewlines("a\r\nb", "\n"), "a\nb");
  });
  test("convierte \\r suelto a \\n", () => {
    assert.equal(normalizeNewlines("a\rb", "\n"), "a\nb");
  });
  test("convierte a \\r\\n cuando se pide", () => {
    assert.equal(normalizeNewlines("a\nb", "\r\n"), "a\r\nb");
  });
  test("no toca texto sin saltos", () => {
    assert.equal(normalizeNewlines("sin saltos", "\n"), "sin saltos");
  });
});

// ---------------------------------------------------------------------------
// parseJsonFromStdout
// ---------------------------------------------------------------------------
describe("parseJsonFromStdout", () => {
  test("parsea JSON aunque haya warnings antes", () => {
    const parsed = parseJsonFromStdout("WARN: múltiples BDs detectadas\n{\"ok\":true}");
    assert.deepEqual(parsed, { ok: true });
  });

  test("parsea JSON pretty y descarta texto posterior", () => {
    const parsed = parseJsonFromStdout("aviso\n[\n  {\"name\":\"A\"}\n]\nfin");
    assert.deepEqual(parsed, [{ name: "A" }]);
  });

  test("lanza error claro cuando no hay JSON válido", () => {
    assert.throws(() => parseJsonFromStdout("WARN sin json", "JSON de prueba"), /JSON de prueba/);
  });
});

// ---------------------------------------------------------------------------
// parseVbaArgsJson
// ---------------------------------------------------------------------------
describe("parseVbaArgsJson", () => {
  test("parsea array de argumentos simples", () => {
    assert.deepEqual(parseVbaArgsJson('[123,"texto",true,null]'), [123, "texto", true, null]);
  });

  test("sin args devuelve array vacío", () => {
    assert.deepEqual(parseVbaArgsJson(""), []);
    assert.deepEqual(parseVbaArgsJson(undefined), []);
  });

  test("rechaza objetos complejos", () => {
    assert.throws(() => parseVbaArgsJson('[{"x":1}]'), /valores simples/);
  });

  test("rechaza JSON que no sea array", () => {
    assert.throws(() => parseVbaArgsJson('"texto"'), /array JSON/);
  });
});

// ---------------------------------------------------------------------------
// evaluateVbaTestExpectation
// ---------------------------------------------------------------------------
describe("evaluateVbaTestExpectation", () => {
  test("valida ok y value desde payload", () => {
    const result = { ok: true, payload: { value: 42 }, returnValue: "ignored", error: null };
    assert.deepEqual(evaluateVbaTestExpectation(result, { ok: true, value: 42 }), { ok: true, failures: [] });
  });

  test("verifica ok: true de forma implícita si se omite en expect", () => {
    const result = { ok: false, payload: null, error: "fallo inesperado" };
    const evaluated = evaluateVbaTestExpectation(result, { value: 42 }); // ok se omite
    assert.equal(evaluated.ok, false);
    assert.ok(evaluated.failures.some(f => f.includes("ok")));
  });

  test("detecta payloadContains fallido", () => {
    const result = { ok: true, payload: { value: 41 }, error: null };
    const evaluated = evaluateVbaTestExpectation(result, { payloadContains: { value: 42 } });
    assert.equal(evaluated.ok, false);
    assert.ok(evaluated.failures[0].includes("payload"));
  });

  test("permite errorContains para tests negativos", () => {
    const result = { ok: false, payload: null, error: "cliente no existe" };
    const evaluated = evaluateVbaTestExpectation(result, { ok: false, errorContains: "no existe" });
    assert.equal(evaluated.ok, true);
  });

  test("evalúa pathEquals correctamente", () => {
    const result = { ok: true, payload: { data: { id: 10 } }, error: null };
    const evaluated = evaluateVbaTestExpectation(result, { pathEquals: { "payload.data.id": 10 } });
    assert.equal(evaluated.ok, true);
    
    const evaluatedFail = evaluateVbaTestExpectation(result, { pathEquals: { "payload.data.id": 99 } });
    assert.equal(evaluatedFail.ok, false);
    assert.ok(evaluatedFail.failures[0].includes("payload.data.id"));
  });
});

// ---------------------------------------------------------------------------
// parseArgs
// ---------------------------------------------------------------------------
describe("parseArgs", () => {
  test("soporta --flag=value", () => {
    const parsed = parseArgs(["node", "cli.js", "start", "--access=MiBD.accdb", "--json"]);
    assert.equal(parsed.command, "start");
    assert.equal(parsed.flags.access, "MiBD.accdb");
    assert.equal(parsed.flags.json, true);
  });
});

// ---------------------------------------------------------------------------
// splitVbaMetadataHeaderText
// ---------------------------------------------------------------------------
describe("splitVbaMetadataHeaderText", () => {
  test("separa atributos/directivas iniciales del cuerpo", () => {
    const parsed = splitVbaMetadataHeaderText(
      "VERSION 1.0 CLASS\r\nAttribute VB_Name = \"Clase1\"\r\nOption Compare Database\r\nOption Explicit\r\n\r\nPublic Sub Foo()\r\nEnd Sub\r\n"
    );

    assert.ok(parsed.header.includes("Attribute VB_Name"));
    assert.ok(parsed.header.includes("Option Explicit"));
    assert.ok(!parsed.header.includes("Public Sub Foo"));
    assert.ok(parsed.body.includes("Public Sub Foo"));
  });

  test("elimina BOM antes de evaluar el header", () => {
    const parsed = splitVbaMetadataHeaderText("\ufeffAttribute VB_Name = \"M\"\nSub Foo()\nEnd Sub\n");

    assert.equal(parsed.header, "Attribute VB_Name = \"M\"");
    assert.equal(parsed.body, "Sub Foo()\nEnd Sub\n");
  });

  test("sin header devuelve cuerpo completo", () => {
    const parsed = splitVbaMetadataHeaderText("Sub Foo()\nEnd Sub\n");

    assert.equal(parsed.header, "");
    assert.equal(parsed.body, "Sub Foo()\nEnd Sub\n");
  });
});
