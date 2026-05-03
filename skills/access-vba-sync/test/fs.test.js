"use strict";

const { test, describe, before, after } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const fsp = fs.promises;
const path = require("path");
const os = require("os");
const { AccessVbaSyncSkill } = require("../handler");

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

async function withTempDir(fn) {
  const dir = await fsp.mkdtemp(path.join(os.tmpdir(), "access-vba-test-"));
  try {
    await fn(dir);
  } finally {
    await fsp.rm(dir, { recursive: true, force: true });
  }
}

function makeSkill(projectRoot, opts = {}) {
  return new AccessVbaSyncSkill({
    skillDir: path.join(__dirname, ".."),
    projectRoot,
    destinationRoot: opts.destinationRoot || "src",
  });
}

async function writeForm(dir, name, { ui = "Begin Form\n  Width=1000\nEnd\n", code = "Option Explicit\nSub Foo()\nEnd Sub\n" } = {}) {
  const formsDir = path.join(dir, "src", "forms");
  await fsp.mkdir(formsDir, { recursive: true });
  const formContent = `Version =21\n${ui}CodeBehind\n${code}`;
  await fsp.writeFile(path.join(formsDir, `${name}.form.txt`), formContent, "utf8");
  return { formsDir, formContent };
}

async function writeCls(dir, name, code = "Option Explicit\nSub Foo()\nEnd Sub\n") {
  const formsDir = path.join(dir, "src", "forms");
  await fsp.mkdir(formsDir, { recursive: true });
  await fsp.writeFile(path.join(formsDir, `${name}.cls`), code, "utf8");
}

// ---------------------------------------------------------------------------
// backupSrcIfExists
// ---------------------------------------------------------------------------
describe("backupSrcIfExists", () => {
  test("no hace nada cuando src no existe", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      await skill.backupSrcIfExists();
      assert.ok(!fs.existsSync(path.join(dir, "src.bak")));
    });
  });

  test("no hace nada cuando src existe pero está vacío", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      await fsp.mkdir(path.join(dir, "src"));
      await skill.backupSrcIfExists();
      assert.ok(!fs.existsSync(path.join(dir, "src.bak")));
    });
  });

  test("crea backup cuando src tiene archivos", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const srcDir = path.join(dir, "src");
      await fsp.mkdir(srcDir);
      await fsp.writeFile(path.join(srcDir, "Mod.bas"), "Sub Foo()\nEnd Sub\n");
      await skill.backupSrcIfExists();
      assert.ok(fs.existsSync(path.join(dir, "src.bak")));
      assert.ok(fs.existsSync(path.join(dir, "src.bak", "Mod.bas")));
    });
  });

  test("el backup preserva el contenido exacto del archivo", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const srcDir = path.join(dir, "src");
      await fsp.mkdir(srcDir);
      const content = "Option Explicit\nSub Calcular(x As Long)\n  x = x + 1\nEnd Sub\n";
      await fsp.writeFile(path.join(srcDir, "Mod.bas"), content);
      await skill.backupSrcIfExists();
      const backed = await fsp.readFile(path.join(dir, "src.bak", "Mod.bas"), "utf8");
      assert.equal(backed, content);
    });
  });

  test("copia subdirectorios recursivamente", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const formsDir = path.join(dir, "src", "forms");
      await fsp.mkdir(formsDir, { recursive: true });
      await fsp.writeFile(path.join(formsDir, "Form_A.form.txt"), "UI");
      await skill.backupSrcIfExists();
      assert.ok(fs.existsSync(path.join(dir, "src.bak", "forms", "Form_A.form.txt")));
    });
  });

  test("sobreescribe el backup anterior", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const srcDir = path.join(dir, "src");
      await fsp.mkdir(srcDir);

      await fsp.writeFile(path.join(srcDir, "v1.bas"), "version 1");
      await skill.backupSrcIfExists();
      assert.ok(fs.existsSync(path.join(dir, "src.bak", "v1.bas")));

      await fsp.unlink(path.join(srcDir, "v1.bas"));
      await fsp.writeFile(path.join(srcDir, "v2.bas"), "version 2");
      await skill.backupSrcIfExists();

      assert.ok(fs.existsSync(path.join(dir, "src.bak", "v2.bas")));
      assert.ok(!fs.existsSync(path.join(dir, "src.bak", "v1.bas")));
    });
  });
});

// ---------------------------------------------------------------------------
// _allDocumentModuleNames
// ---------------------------------------------------------------------------
describe("_allDocumentModuleNames", () => {
  test("encuentra todos los .form.txt en src/forms", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const formsDir = path.join(dir, "src", "forms");
      await fsp.mkdir(formsDir, { recursive: true });
      await fsp.writeFile(path.join(formsDir, "Form_A.form.txt"), "");
      await fsp.writeFile(path.join(formsDir, "Form_B.form.txt"), "");
      await fsp.writeFile(path.join(formsDir, "ignore.cls"), "");

      const names = await skill._allDocumentModuleNames();
      assert.ok(names.includes("Form_A"));
      assert.ok(names.includes("Form_B"));
      assert.ok(!names.includes("ignore"));
      assert.equal(names.length, 2);
    });
  });

  test("encuentra .report.txt en src/reports", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const reportsDir = path.join(dir, "src", "reports");
      await fsp.mkdir(reportsDir, { recursive: true });
      await fsp.writeFile(path.join(reportsDir, "Report_X.report.txt"), "");

      const names = await skill._allDocumentModuleNames();
      assert.ok(names.includes("Report_X"));
    });
  });

  test("devuelve array vacío cuando no existe src/forms ni src/reports", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      assert.deepEqual(await skill._allDocumentModuleNames(), []);
    });
  });
});

// ---------------------------------------------------------------------------
// importModules
// ---------------------------------------------------------------------------
describe("importModules", () => {
  test("fallback módulo-a-módulo registra y ejecuta cada import individual", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const dbPath = path.join(dir, "TEST.accdb");
      const srcPath = path.join(dir, "src");
      await fsp.mkdir(srcPath, { recursive: true });
      await fsp.writeFile(dbPath, "fake db");

      skill.resolveCommandContext = async () => ({
        accessPath: dbPath,
        destinationRootAbs: srcPath,
        modulesPath: srcPath,
      });
      skill.syncCodeBehind = async () => 0;

      const calls = [];
      skill.runVbaManager = async (args) => {
        calls.push(args);
        if (calls.length === 1) {
          const err = new Error("array binding failed");
          err.stderr = "ParameterBindingException: cannot process argument transformation";
          err.stdout = "";
          throw err;
        }
        return {
          code: 0,
          stdout: `##MODULE_RESULTS:[{"module":"${args.moduleNames[0]}","status":"ok","message":""}]\n`,
          stderr: "",
        };
      };

      await skill.importModules({ moduleNames: ["ModuloA", "ModuloB"], importMode: "Auto" });

      assert.equal(calls.length, 3);
      assert.deepEqual(calls[0].moduleNames, ["ModuloA", "ModuloB"]);
      assert.deepEqual(calls[1].moduleNames, ["ModuloA"]);
      assert.deepEqual(calls[2].moduleNames, ["ModuloB"]);
    });
  });

  test("fallback módulo-a-módulo continúa y reporta fallos al final", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const dbPath = path.join(dir, "TEST.accdb");
      const srcPath = path.join(dir, "src");
      await fsp.mkdir(srcPath, { recursive: true });
      await fsp.writeFile(dbPath, "fake db");

      skill.resolveCommandContext = async () => ({
        accessPath: dbPath,
        destinationRootAbs: srcPath,
        modulesPath: srcPath,
      });
      skill.syncCodeBehind = async () => 0;

      const calls = [];
      skill.runVbaManager = async (args) => {
        calls.push(args);
        if (calls.length === 1) {
          const err = new Error("array binding failed");
          err.stderr = "ParameterBindingException: cannot process argument transformation";
          err.stdout = "";
          throw err;
        }
        if (args.moduleNames[0] === "ModuloA") {
          const err = new Error("falló ModuloA");
          err.stdout = '##MODULE_RESULTS:[{"module":"ModuloA","status":"error","error":"falló ModuloA"}]';
          err.stderr = "";
          throw err;
        }
        return {
          code: 0,
          stdout: `##MODULE_RESULTS:[{"module":"${args.moduleNames[0]}","status":"ok","message":""}]\n`,
          stderr: "",
        };
      };

      await assert.rejects(
        () => skill.importModules({ moduleNames: ["ModuloA", "ModuloB", "ModuloC"], importMode: "Auto" }),
        /Falló el import individual de 1\/3/
      );

      assert.equal(calls.length, 4);
      assert.deepEqual(calls[1].moduleNames, ["ModuloA"]);
      assert.deepEqual(calls[2].moduleNames, ["ModuloB"]);
      assert.deepEqual(calls[3].moduleNames, ["ModuloC"]);
    });
  });
});

// ---------------------------------------------------------------------------
// importAll
// ---------------------------------------------------------------------------
describe("importAll", () => {
  test("restaura el binario desde .bak si falla el import completo", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const dbPath = path.join(dir, "TEST.accdb");
      const srcPath = path.join(dir, "src");
      await fsp.mkdir(srcPath, { recursive: true });
      await fsp.writeFile(dbPath, "original");

      skill.resolveCommandContext = async () => ({
        accessPath: dbPath,
        destinationRootAbs: srcPath,
        modulesPath: srcPath,
      });
      skill.syncAllDocumentCodeBehind = async () => 0;
      skill.runVbaManager = async () => {
        await fsp.writeFile(dbPath, "corrupted");
        throw new Error("import failed");
      };

      await assert.rejects(() => skill.importAll({}), /import failed/);
      assert.equal(await fsp.readFile(dbPath, "utf8"), "original");
      assert.ok(!fs.existsSync(dbPath + ".bak"));
    });
  });
});

// ---------------------------------------------------------------------------
// generateErd
// ---------------------------------------------------------------------------
describe("generateErd", () => {
  test("no pasa AccessPath vacío al VBAManager", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      let captured;
      skill.runVbaManager = async (args) => {
        captured = args;
        return { code: 0, stdout: "", stderr: "" };
      };

      await skill.generateErd({ backendPath: path.join(dir, "Datos.accdb") });

      assert.equal(captured.action, "Generate-ERD");
      assert.equal(Object.prototype.hasOwnProperty.call(captured, "accessPath"), false);
    });
  });
});

// ---------------------------------------------------------------------------
// runVba
// ---------------------------------------------------------------------------
describe("runVba", () => {
  test("ejecuta procedimiento público pasando args-json como array", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const dbPath = path.join(dir, "TEST.accdb");
      const srcPath = path.join(dir, "src");
      await fsp.mkdir(srcPath, { recursive: true });
      await fsp.writeFile(dbPath, "fake db");

      skill.resolveCommandContext = async () => ({
        accessPath: dbPath,
        destinationRootAbs: srcPath,
        modulesPath: srcPath,
      });

      let received;
      skill.runVbaManager = async (args) => {
        received = args;
        return {
          code: 0,
          stdout: JSON.stringify({
            ok: true,
            procedure: args.procedureName,
            argsCount: args.procedureArgs.length,
            returnValue: '{"ok":true,"logs":["preparando","resultado OK"],"value":"OK"}',
            returnType: "System.String",
            payload: { ok: true, logs: ["preparando", "resultado OK"], value: "OK" },
            logs: ["preparando", "resultado OK"],
            error: null,
          }),
          stderr: "",
        };
      };

      const result = await skill.runVba({
        procedureName: "Test_CalculaTotal",
        argsJson: '[123,"abc",true]',
        accessPath: dbPath,
        json: true,
      });

      assert.equal(received.action, "Run-Procedure");
      assert.equal(received.procedureName, "Test_CalculaTotal");
      assert.deepEqual(received.procedureArgs, [123, "abc", true]);
      assert.equal(received.json, true);
      assert.equal(result.ok, true);
      assert.deepEqual(result.logs, ["preparando", "resultado OK"]);
      assert.equal(result.payload.value, "OK");
    });
  });
});

// ---------------------------------------------------------------------------
// compileVba
// ---------------------------------------------------------------------------
describe("compileVba", () => {
  test("devuelve resultado JSON de compilación", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const dbPath = path.join(dir, "TEST.accdb");
      const srcPath = path.join(dir, "src");
      await fsp.mkdir(srcPath, { recursive: true });
      await fsp.writeFile(dbPath, "fake db");

      skill.resolveCommandContext = async () => ({
        accessPath: dbPath,
        destinationRootAbs: srcPath,
        modulesPath: srcPath,
      });
      skill.runVbaManager = async (args) => ({
        code: 0,
        stdout: JSON.stringify({ ok: true, phase: "compile", error: null }),
        stderr: "",
      });

      const result = await skill.compileVba({ accessPath: dbPath, json: true });
      assert.equal(result.ok, true);
      assert.equal(result.phase, "compile");
    });
  });
});

// ---------------------------------------------------------------------------
// testVba
// ---------------------------------------------------------------------------
describe("testVba", () => {
  test("compila y ejecuta plan de tests externo", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const dbPath = path.join(dir, "TEST.accdb");
      const srcPath = path.join(dir, "src");
      await fsp.mkdir(srcPath, { recursive: true });
      await fsp.writeFile(dbPath, "fake db");
      await fsp.writeFile(
        path.join(dir, "tests.vba.json"),
        JSON.stringify({
          tests: [
            { name: "total básico", procedure: "Test_Total", args: [10], expect: { ok: true, value: 42 } },
            { name: "cliente negativo", procedure: "Test_Cliente", args: ["X"], expect: { ok: false, errorContains: "no existe" } },
          ],
        }),
        "utf8"
      );

      skill.resolveCommandContext = async () => ({
        accessPath: dbPath,
        destinationRootAbs: srcPath,
        modulesPath: srcPath,
      });
      skill.compileVba = async () => ({ ok: true, phase: "compile" });
      skill.runVba = async ({ procedureName }) => {
        if (procedureName === "Test_Total") {
          return { ok: true, procedure: procedureName, payload: { value: 42 }, logs: ["ok"], error: null };
        }
        return { ok: false, procedure: procedureName, payload: null, logs: ["negativo"], error: "cliente no existe" };
      };

      const report = await skill.testVba({ testsPath: "tests.vba.json", accessPath: dbPath, json: true });
      assert.equal(report.ok, true);
      assert.equal(report.total, 2);
      assert.equal(report.passed, 2);
      assert.equal(report.failed, 0);
      assert.deepEqual(report.results[0].logs, ["ok"]);
    });
  });

  test("no ejecuta tests si falla compile gate", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const dbPath = path.join(dir, "TEST.accdb");
      const srcPath = path.join(dir, "src");
      await fsp.mkdir(srcPath, { recursive: true });
      await fsp.writeFile(dbPath, "fake db");
      await fsp.writeFile(
        path.join(dir, "tests.vba.json"),
        JSON.stringify({ tests: [{ name: "no corre", procedure: "Test_NoCorre" }] }),
        "utf8"
      );

      let ran = false;
      skill.compileVba = async () => ({
        ok: false,
        phase: "compile",
        component: "ModuloA",
        line: 12,
        sourceLine: "Call Foo()",
        error: "Argument not optional",
      });
      skill.runVba = async () => {
        ran = true;
      };

      const report = await skill.testVba({ testsPath: "tests.vba.json", accessPath: dbPath, json: true });
      assert.equal(report.ok, false);
      assert.equal(report.phase, "compile");
      assert.equal(report.skipped, 1);
      assert.equal(ran, false);
    });
  });

  test("ejecuta un procedimiento directo sin tests.vba.json", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const dbPath = path.join(dir, "TEST.accdb");
      const srcPath = path.join(dir, "src");
      await fsp.mkdir(srcPath, { recursive: true });
      await fsp.writeFile(dbPath, "fake db");

      skill.compileVba = async () => ({ ok: true, phase: "compile" });
      let received;
      skill.runVba = async (args) => {
        received = args;
        return { ok: true, procedure: args.procedureName, payload: null, logs: ["canonical ok"], error: null };
      };

      const report = await skill.testVba({
        procedureName: "Canonical_RunAll",
        argsJson: '["DC"]',
        accessPath: dbPath,
        json: true
      });

      assert.equal(received.procedureName, "Canonical_RunAll");
      assert.equal(received.argsJson, "[\"DC\"]");
      assert.equal(report.ok, true);
      assert.equal(report.testsPath, null);
      assert.equal(report.total, 1);
      assert.equal(report.passed, 1);
      assert.deepEqual(report.results[0].logs, ["canonical ok"]);
    });
  });

  test("filter sin coincidencias devuelve ok:true con total:0 y skipped igual al total del plan", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const dbPath = path.join(dir, "TEST.accdb");
      const srcPath = path.join(dir, "src");
      await fsp.mkdir(srcPath, { recursive: true });
      await fsp.writeFile(dbPath, "fake db");
      await fsp.writeFile(
        path.join(dir, "tests.vba.json"),
        JSON.stringify({
          tests: [
            { name: "test A", procedure: "Test_A", tags: ["area1"] },
            { name: "test B", procedure: "Test_B", tags: ["area2"] },
          ],
        }),
        "utf8"
      );

      skill.compileVba = async () => ({ ok: true, phase: "compile" });
      let ran = false;
      skill.runVba = async () => { ran = true; };

      const report = await skill.testVba({
        testsPath: "tests.vba.json",
        filter: "noexiste",
        accessPath: dbPath,
        json: true,
      });

      assert.equal(ran, false);
      assert.equal(report.ok, true);
      assert.equal(report.total, 0);
      assert.equal(report.passed, 0);
      assert.equal(report.skipped, 2);
    });
  });
});

// ---------------------------------------------------------------------------
// syncCodeBehind
// ---------------------------------------------------------------------------
describe("syncCodeBehind", () => {
  test("actualiza el CodeBehind del .form.txt con el contenido del .cls", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      await writeForm(dir, "Form_Test", { code: "Sub OldSub()\nEnd Sub\n" });
      await writeCls(dir, "Form_Test", "Option Explicit\nSub NewSub()\nEnd Sub\n");

      await skill.syncCodeBehind(["Form_Test"]);

      const updated = fs.readFileSync(
        path.join(dir, "src", "forms", "Form_Test.form.txt"),
        "utf8"
      );
      assert.ok(updated.includes("NewSub"), "debería incluir el nuevo código");
      assert.ok(!updated.includes("OldSub"), "no debería tener el código viejo");
      assert.ok(updated.includes("Begin Form"), "debe preservar la sección UI");
    });
  });

  test("no modifica el .form.txt cuando no hay .cls", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      await writeForm(dir, "Form_NoCode", { code: "Sub Original()\nEnd Sub\n" });

      await skill.syncCodeBehind(["Form_NoCode"]);

      const content = fs.readFileSync(
        path.join(dir, "src", "forms", "Form_NoCode.form.txt"),
        "utf8"
      );
      assert.ok(content.includes("Original"), "el contenido debe permanecer igual");
    });
  });

  test("devuelve el número de documentos sincronizados", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      await writeForm(dir, "Form_A");
      await writeCls(dir, "Form_A", "Sub NuevoA()\nEnd Sub\n");
      await writeForm(dir, "Form_B");
      // Form_B no tiene .cls

      const count = await skill.syncCodeBehind(["Form_A", "Form_B"]);
      assert.equal(count, 1);
    });
  });

  test("advierte cuando falla la unificación del sidecar espejo", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      await writeForm(dir, "Form_Test", { code: "Sub OldSub()\nEnd Sub\n" });
      const clsPath = path.join(dir, "src", "forms", "Form_Test.cls");
      await writeCls(dir, "Form_Test", "Option Explicit\nSub NewSub()\nEnd Sub\n");

      skill.resolveDocumentArtifacts = () => ({
        moduleName: "Form_Test",
        textPath: path.join(dir, "src", "forms", "Form_Test.form.txt"),
        clsPath,
        mirrorClsPaths: [path.join(dir, "src", "forms", "no-existe", "Form_Test.cls")],
      });

      const warnings = [];
      const originalWarn = console.warn;
      console.warn = (msg) => warnings.push(String(msg));
      try {
        await skill.syncCodeBehind(["Form_Test"]);
      } finally {
        console.warn = originalWarn;
      }

      assert.ok(warnings.some((msg) => msg.includes("no se pudo unificar sidecar")));
    });
  });
});

// ---------------------------------------------------------------------------
// verifyCode
// ---------------------------------------------------------------------------
describe("verifyCode", () => {
  test("reporta in-sync cuando .cls y CodeBehind coinciden", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const code = "Option Explicit\nSub Foo()\nEnd Sub\n";
      await writeForm(dir, "Form_Test", { code });
      await writeCls(dir, "Form_Test", code);

      const inSync = await skill.verifyCode({ moduleNames: ["Form_Test"] });
      assert.ok(inSync);
    });
  });

  test("detecta desincronización entre .cls y CodeBehind", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      await writeForm(dir, "Form_Test", { code: "Sub OldSub()\nEnd Sub\n" });
      await writeCls(dir, "Form_Test", "Sub NewSub()\nEnd Sub\n");

      const inSync = await skill.verifyCode({ moduleNames: ["Form_Test"] });
      assert.ok(!inSync);
    });
  });

  test("sin argumentos verifica todos los formularios encontrados", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const code = "Sub Foo()\nEnd Sub\n";
      await writeForm(dir, "Form_A", { code });
      await writeCls(dir, "Form_A", code);
      await writeForm(dir, "Form_B", { code });
      await writeCls(dir, "Form_B", code);

      const inSync = await skill.verifyCode({});
      assert.ok(inSync);
    });
  });

  test("devuelve true cuando no hay formularios", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const inSync = await skill.verifyCode({});
      assert.ok(inSync);
    });
  });

  test("formulario sin .cls no cuenta como error de sync", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      await writeForm(dir, "Form_SinCls");

      const inSync = await skill.verifyCode({ moduleNames: ["Form_SinCls"] });
      assert.ok(inSync);
    });
  });
});

// ---------------------------------------------------------------------------
// resolveDocumentArtifacts
// ---------------------------------------------------------------------------
describe("resolveDocumentArtifacts", () => {
  test("resuelve Form_X desde nombre completo", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      await writeForm(dir, "Form_frmGestion");
      await writeCls(dir, "Form_frmGestion");

      const artifacts = skill.resolveDocumentArtifacts("Form_frmGestion");
      assert.ok(artifacts !== null);
      assert.ok(artifacts.textPath.endsWith("Form_frmGestion.form.txt"));
    });
  });

  test("resuelve desde nombre sin prefijo (frmGestion → Form_frmGestion)", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      await writeForm(dir, "Form_frmGestion");

      const artifacts = skill.resolveDocumentArtifacts("frmGestion");
      assert.ok(artifacts !== null);
    });
  });

  test("devuelve null cuando no existe el .form.txt", async () => {
    await withTempDir(async (dir) => {
      const skill = makeSkill(dir);
      const artifacts = skill.resolveDocumentArtifacts("Form_Inexistente");
      assert.equal(artifacts, null);
    });
  });
});
