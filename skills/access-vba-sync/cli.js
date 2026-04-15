#!/usr/bin/env node
"use strict";

const path = require("path");
const { AccessVbaSyncSkill } = require("./handler");

function parseArgs(argv) {
  const args = argv.slice(2);
  const out = {
    command: null,
    mods: [],
    flags: {}
  };

  if (args.length === 0) return out;
  out.command = args[0];

  for (let i = 1; i < args.length; i++) {
    const a = args[i];
    if (a.startsWith("--")) {
      const key = a.slice(2);
      const val = args[i + 1];
      if (val == null || val.startsWith("--")) {
        out.flags[key] = true;
      } else {
        out.flags[key] = val;
        i++;
      }
      continue;
    }
    out.mods.push(a);
  }

  return out;
}

function toBoolFlag(v, defaultValue) {
  if (v === undefined) return defaultValue;
  if (typeof v === "boolean") return v;
  const s = String(v).trim().toLowerCase();
  if (s === "false" || s === "0" || s === "no") return false;
  if (s === "true" || s === "1" || s === "yes") return true;
  return defaultValue;
}

function normalizePathFlag(p) {
  if (!p) return null;
  return path.resolve(process.cwd(), p);
}

function printHelp() {
  console.log(
    [
      "Uso:",
      "  node cli.js start          [--access <ruta>] [--destination_root <dir>] [--password <pwd>]",
      "  node cli.js watch          [--access <ruta>] [--destination_root <dir>] [--debounce_ms <n>] [--password <pwd>]",
      "  node cli.js export  <Mod..>[--access <ruta>] [--destination_root <dir>] [--password <pwd>]",
      "  node cli.js export-all     [--access <ruta>] [--destination_root <dir>] [--password <pwd>]",
      "  node cli.js import  <Mod..>[--access <ruta>] [--destination_root <dir>] [--password <pwd>]",
      "  node cli.js import-form <Mod..>[--access <ruta>] [--destination_root <dir>] [--password <pwd>]",
      "  node cli.js import-code <Mod..>[--access <ruta>] [--destination_root <dir>] [--password <pwd>]",
      "  node cli.js import-all     [--access <ruta>] [--destination_root <dir>] [--password <pwd>]",
      "  node cli.js sync    <Mod..>[--access <ruta>] [--destination_root <dir>] [--password <pwd>]",
      "  node cli.js fix-encoding   [--access <ruta>] [--destination_root <dir>] [--password <pwd>] [--location Both|Src|Access] [<Mod...>]",
      "  node cli.js generate-erd   [--backend <ruta>] [--erd_path <dir>] [--password <pwd>]",
      "  node cli.js sandbox        [--access <ruta>] [--password <pwd>] [--backend_password <pwd>] [--keep_sidecars]",
      "  node cli.js status",
      "  node cli.js end            [--auto_export_on_end false]",
      "",
      "Comandos:",
      "  start            Export inicial de todos los módulos + inicia sesión",
      "  watch            Inicia sesión (si no hay) + auto-sync al guardar archivos",
      "  export  <Mod..>  Exporta módulos específicos de la BD hacia src/",
      "  export-all       Exporta todos los módulos de la BD hacia src/",
      "  import  <Mod..>  Importa módulos específicos de src/ hacia la BD",
      "  import-form <Mod..> Importa formularios desde *.form.txt (UI + código)",
      "  import-code <Mod..> Importa code-behind desde *.cls/*.bas (sin layout)",
      "  import-all       Importa todos los módulos de src/ hacia la BD",
      "  sync    <Mod..>  Alias de import",
      "  fix-encoding     Corrige encoding (ANSI→UTF-8 sin BOM) en src/, en la BD, o en ambos",
      "                   Sin módulos: procesa todos. Con módulos: solo los indicados.",
      "  generate-erd     Genera documentación de estructura de tablas en Markdown",
      "  sandbox          Copia backends vinculados al lado del frontend y revincula las tablas",
      "                   creando un sandbox aislado de producción",
      "  status           Muestra el estado de la sesión activa",
      "  end              Cierra la sesión y restaura la configuración de Access",
      "",
      "Flags comunes:",
      "  --access <ruta>              Ruta .accdb/.accde/.mdb/.mde (relativa a CWD o absoluta)",
      "  --password <pwd>             Contraseña de la BD si está protegida",
      "  --destination_root <dir>     Carpeta de export/import (default: src)",
      "",
      "Flags específicos:",
      "  --debounce_ms <n>            Debounce para watch en ms (default: 600)",
      "  --auto_export_on_end false   Desactiva export final al cerrar sesión",
      "  --location Both|Src|Access   Para fix-encoding: dónde aplicar (default: Both)",
      "  --backend <ruta>             Para generate-erd: ruta al backend _Datos.accdb",
      "  --erd_path <dir>             Para generate-erd: carpeta de salida (default: docs/ERD)",
      "  --backend_password <pwd>     Para sandbox: contraseña de los backends (default: misma que --password)",
      "  --keep_sidecars              Para sandbox: no borrar los backends copiados al terminar"
    ].join("\n")
  );
}

async function main() {
  const { command, mods, flags } = parseArgs(process.argv);
  if (!command || command === "help" || command === "--help" || command === "-h") {
    printHelp();
    process.exitCode = command ? 0 : 1;
    return;
  }

  const skill = new AccessVbaSyncSkill({
    skillDir: __dirname,
    projectRoot: process.cwd(),
    destinationRoot: flags.destination_root || "src",
    debounceMs: Number.isFinite(Number(flags.debounce_ms)) ? Number(flags.debounce_ms) : 600,
    autoExportOnEnd: toBoolFlag(flags.auto_export_on_end, true),
    password: flags.password || null
  });

  const accessPath = normalizePathFlag(flags.access);

  if (command === "start") {
    await skill.start({ accessPath });
    return;
  }

  if (command === "watch") {
    await skill.watch({ accessPath });
    return;
  }

  if (command === "export") {
    if (mods.length === 0) {
      console.error("Faltan módulos. Ejemplo: node cli.js export Form_FormInicial Utilidades");
      process.exitCode = 1;
      return;
    }
    await skill.exportModules({ moduleNames: mods, accessPath });
    return;
  }

  if (command === "export-all") {
    await skill.exportAll({ accessPath });
    return;
  }

  if (command === "import" || command === "sync") {
    if (mods.length === 0) {
      console.error("Faltan módulos. Ejemplo: node cli.js import Utilidades Validaciones");
      process.exitCode = 1;
      return;
    }
    await skill.importModules({ moduleNames: mods, accessPath });
    return;
  }

  if (command === "import-form") {
    if (mods.length === 0) {
      console.error("Faltan módulos. Ejemplo: node cli.js import-form Form_frmDatosPC");
      process.exitCode = 1;
      return;
    }
    await skill.importForms({ moduleNames: mods, accessPath });
    return;
  }

  if (command === "import-code") {
    if (mods.length === 0) {
      console.error("Faltan módulos. Ejemplo: node cli.js import-code Form_frmDatosPC");
      process.exitCode = 1;
      return;
    }
    await skill.importCode({ moduleNames: mods, accessPath });
    return;
  }

  if (command === "import-all") {
    await skill.importAll({ accessPath });
    return;
  }

  if (command === "fix-encoding") {
    const location = flags.location || "Both";
    const validLocations = ["Both", "Src", "Access"];
    if (!validLocations.includes(location)) {
      console.error(`--location debe ser uno de: ${validLocations.join(", ")}`);
      process.exitCode = 1;
      return;
    }
    await skill.fixEncoding({ moduleNames: mods.length > 0 ? mods : null, accessPath, location });
    return;
  }

  if (command === "generate-erd") {
    const backendPath = normalizePathFlag(flags.backend);
    const erdPath = normalizePathFlag(flags.erd_path || "docs/ERD");
    await skill.generateErd({ backendPath, erdPath });
    return;
  }

  if (command === "sandbox") {
    const accessPath = normalizePathFlag(flags.access);
    const backendPassword = flags.backend_password || null;
    const keepSidecars = toBoolFlag(flags.keep_sidecars, false);
    await skill.sandbox({ accessPath, backendPassword, keepSidecars });
    return;
  }

  if (command === "status") {
    await skill.status();
    return;
  }

  if (command === "end") {
    await skill.end();
    return;
  }

  console.error(`Comando desconocido: ${command}`);
  printHelp();
  process.exitCode = 1;
}

main().catch((err) => {
  console.error("ERROR:", err && err.message ? err.message : String(err));
  if (err && err.stdout) {
    console.error("\n--- stdout ---\n" + err.stdout);
  }
  if (err && err.stderr) {
    console.error("\n--- stderr ---\n" + err.stderr);
  }
  process.exitCode = 1;
});
