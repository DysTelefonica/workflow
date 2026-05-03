#!/usr/bin/env node
"use strict";

/**
 * Smoke test manual — requiere Microsoft Access instalado y una BD real.
 *
 * Uso:
 *   node test/smoke.js --access "C:\ruta\a\MiBD.accdb" [--password <pwd>]
 *
 * Ejecuta los casos de SKILL.md "Pruebas mínimas" de forma headless.
 * Crea un src/ temporal dentro de un directorio aislado para no tocar el proyecto.
 */

const path = require("path");
const fs = require("fs");
const fsp = fs.promises;
const os = require("os");
const { spawn } = require("child_process");

// ---------------------------------------------------------------------------
// CLI args
// ---------------------------------------------------------------------------
const args = process.argv.slice(2);
let accessPath = null;
let password = null;

for (let i = 0; i < args.length; i++) {
  if (args[i] === "--access" && args[i + 1]) { accessPath = args[++i]; continue; }
  if (args[i] === "--password" && args[i + 1]) { password = args[++i]; continue; }
}

if (!accessPath) {
  console.error("Uso: node test/smoke.js --access <ruta.accdb> [--password <pwd>]");
  process.exit(1);
}

const CLI = path.join(__dirname, "..", "cli.js");

// ---------------------------------------------------------------------------
// Runner
// ---------------------------------------------------------------------------
let passed = 0;
let failed = 0;

async function run(label, fn) {
  process.stdout.write(`  ${label} ... `);
  try {
    await fn();
    console.log("✅ OK");
    passed++;
  } catch (err) {
    console.log(`❌ FAIL\n     ${err.message || String(err)}`);
    failed++;
  }
}

function cli(cwd, extraArgs) {
  return new Promise((resolve, reject) => {
    const nodeArgs = [CLI, ...extraArgs];
    if (password) nodeArgs.push("--password", password);

    const child = spawn(process.execPath, nodeArgs, {
      cwd,
      windowsHide: true,
      env: { ...process.env }
    });

    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (d) => (stdout += d));
    child.stderr.on("data", (d) => (stderr += d));
    child.on("close", (code) => {
      if (code === 0) return resolve({ stdout, stderr });
      const err = new Error(`Exit ${code}\nstdout: ${stdout}\nstderr: ${stderr}`);
      err.stdout = stdout;
      err.stderr = stderr;
      reject(err);
    });
    child.on("error", reject);
  });
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------
async function main() {
  const workDir = await fsp.mkdtemp(path.join(os.tmpdir(), "access-vba-smoke-"));

  // Copiar la BD al directorio temporal de trabajo
  const dbName = path.basename(accessPath);
  const dbDest = path.join(workDir, dbName);
  await fsp.copyFile(accessPath, dbDest);

  console.log(`\n🔬 Smoke test — BD: ${dbName}`);
  console.log(`   Directorio de trabajo: ${workDir}\n`);

  try {
    // ------------------------------------------------------------------
    // 1. start con BD única → crea src/modules/, src/classes/, src/forms/
    // ------------------------------------------------------------------
    await run("start: crea src/ con subcarpetas", async () => {
      await cli(workDir, ["start", "--access", dbName]);
      const srcDir = path.join(workDir, "src");
      assert(fs.existsSync(srcDir), "src/ no existe");
      const entries = fs.readdirSync(srcDir);
      assert(entries.length > 0, "src/ está vacío tras el export");
    });

    // ------------------------------------------------------------------
    // 2. start sin BD en CWD → error claro
    // ------------------------------------------------------------------
    await run("start: error claro cuando no hay BD en CWD", async () => {
      const emptyDir = await fsp.mkdtemp(path.join(os.tmpdir(), "access-vba-empty-"));
      try {
        await cli(emptyDir, ["start"]);
        throw new Error("Debería haber fallado pero tuvo éxito");
      } catch (err) {
        if (err.message.startsWith("Debería haber fallado")) throw err;
        // Error esperado — OK
      } finally {
        await fsp.rm(emptyDir, { recursive: true, force: true });
      }
    });

    // ------------------------------------------------------------------
    // 3. list-objects → devuelve inventario
    // ------------------------------------------------------------------
    await run("list-objects: devuelve inventario JSON no vacío", async () => {
      const result = await cli(workDir, ["list-objects", "--access", dbName, "--json"]);
      const data = JSON.parse(result.stdout.trim());
      const total = Object.values(data).reduce((s, v) => s + (Array.isArray(v) ? v.length : 0), 0);
      assert(total > 0, "El inventario está vacío");
    });

    // ------------------------------------------------------------------
    // 4. export-all → actualiza src/ (backup + re-export)
    // ------------------------------------------------------------------
    await run("export-all: crea backup y re-exporta", async () => {
      await cli(workDir, ["export-all", "--access", dbName]);
      assert(fs.existsSync(path.join(workDir, "src.bak")), "src.bak no se creó");
    });

    // ------------------------------------------------------------------
    // 5. fix-encoding --location Src → no lanza error
    // ------------------------------------------------------------------
    await run("fix-encoding --location Src: termina sin error", async () => {
      await cli(workDir, ["fix-encoding", "--access", dbName, "--location", "Src"]);
    });

    // ------------------------------------------------------------------
    // 6. import-all → importa todo src/ con backup del binario
    // ------------------------------------------------------------------
    await run("import-all: importa src/ completo y elimina .bak del binario", async () => {
      await cli(workDir, ["import-all", "--access", dbName]);
      // El backup del binario debería haberse borrado tras éxito
      const bakPath = dbDest + ".bak";
      assert(!fs.existsSync(bakPath), "El backup del binario no se eliminó tras éxito");
    });

    // ------------------------------------------------------------------
    // 7. verify-code → sin diferencias tras export+import limpio
    // ------------------------------------------------------------------
    await run("verify-code: sin desincronización tras ciclo export→import", async () => {
      const result = await cli(workDir, ["verify-code"]);
      assert(!result.stdout.includes("DESINCRONIZADO"), "Hay módulos desincronizados tras ciclo limpio");
    });

    // ------------------------------------------------------------------
    // 8. status → sin sesión activa (start-end completo)
    // ------------------------------------------------------------------
    await run("status: refleja estado correcto (sin sesión)", async () => {
      const result = await cli(workDir, ["status"]);
      assert(
        result.stdout.includes("SIN SESIÓN") || result.stdout.includes("Activa"),
        "status no devuelve información reconocible"
      );
    });

  } finally {
    await fsp.rm(workDir, { recursive: true, force: true });
  }

  console.log(`\n${"─".repeat(50)}`);
  console.log(`Resultado: ${passed} passed, ${failed} failed`);
  if (failed > 0) process.exitCode = 1;
}

function assert(condition, message) {
  if (!condition) throw new Error(message || "Assertion failed");
}

main().catch((err) => {
  console.error("\nERROR inesperado en smoke test:", err.message || String(err));
  process.exitCode = 1;
});
