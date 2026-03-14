#!/usr/bin/env node
"use strict"

const fs   = require("fs")
const path = require("path")
const { execSync } = require("child_process")

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function ensureDir(p) {
  if (!fs.existsSync(p)) fs.mkdirSync(p, { recursive: true })
}

function copyDir(src, dest) {
  if (!fs.existsSync(src)) return
  ensureDir(dest)
  for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
    const srcPath  = path.join(src, entry.name)
    const destPath = path.join(dest, entry.name)
    if (entry.isDirectory()) {
      copyDir(srcPath, destPath)
    } else {
      fs.copyFileSync(srcPath, destPath)
    }
  }
}

function log(icon, msg) {
  console.log(`  ${icon} ${msg}`)
}

// ---------------------------------------------------------------------------
// Skills dir detection (sin interacción — solo autodetect)
// ---------------------------------------------------------------------------

function detectSkillsDir(cwd) {
  if (fs.existsSync(path.join(cwd, ".agents", "skills"))) return { skillsDir: ".agents/skills", rulesDir: ".agents/rules" }
  if (fs.existsSync(path.join(cwd, ".agents")))           return { skillsDir: ".agents/skills", rulesDir: ".agents/rules" }
  if (fs.existsSync(path.join(cwd, ".trae", "skills")))  return { skillsDir: ".trae/skills", rulesDir: ".trae/rules" }
  if (fs.existsSync(path.join(cwd, ".agent", "skills"))) return { skillsDir: ".agent/skills", rulesDir: ".agent/rules" }
  if (fs.existsSync(path.join(cwd, "skills")))           return { skillsDir: "skills", rulesDir: "rules" }
  return null
}

// ---------------------------------------------------------------------------
// Main updater
// ---------------------------------------------------------------------------

module.exports = function updateAccess() {
  const cwd     = process.cwd()
  const pkgRoot = path.join(__dirname, "..")
  const pkgVersion = require("../package.json").version

  console.log("\n\n  🔄  DYSFLOW — UPDATE\n" + "  " + "=".repeat(40))

  // 1. Detectar skills dir
  const dirs = detectSkillsDir(cwd)
  if (!dirs) {
    console.error("\n  ✗  No se encontró ninguna carpeta de skills.")
    console.error("     Usa primero: dysflow init access\n")
    process.exit(1)
  }
  const { skillsDir, rulesDir } = dirs
  log("✔", `Skills detectadas en: ${skillsDir}`)

  // 2. Leer versión instalada
  const versionFile = path.join(cwd, skillsDir, ".dysflow")
  let installedVersion = null
  if (fs.existsSync(versionFile)) {
    try {
      installedVersion = JSON.parse(fs.readFileSync(versionFile, "utf8")).version
    } catch (_) {}
  }

  if (installedVersion) {
    log("·", `Versión instalada: ${installedVersion}`)
  } else {
    log("·", "Sin versión registrada (instalación antigua)")
  }
  log("·", `Versión del paquete: ${pkgVersion}`)

  if (installedVersion === pkgVersion) {
    console.log(`\n  Skills ya en v${pkgVersion} — sincronizando archivos de framework...\n`)
  } else if (installedVersion) {
    console.log(`\n  Actualizando ${installedVersion} → ${pkgVersion}...\n`)
  } else {
    console.log(`\n  Sincronizando skills a v${pkgVersion}...\n`)
  }

  // 3. Actualizar solo archivos de framework
  // Skills
  copyDir(path.join(pkgRoot, "skills"), path.join(cwd, skillsDir))
  log("✔", `skills/ actualizado en ${skillsDir}`)

  // Rules
  copyDir(path.join(pkgRoot, "rules"), path.join(cwd, rulesDir))
  log("✔", `rules/ actualizado en ${rulesDir}`)

  // Templates
  if (fs.existsSync(path.join(cwd, "docs", "templates"))) {
    copyDir(path.join(pkgRoot, "templates"), path.join(cwd, "docs", "templates"))
    log("✔", "templates/ actualizado en docs/templates/")
  }

  // 3.5 Bootstrap minimo de estructura (solo crea faltantes, no sobreescribe)
  const requiredDirs = [
    "docs/plans/active",
    "docs/plans/completed",
    "docs/specs/active",
    "docs/specs/completed",
    "docs/PRD",
    "docs/ERD",
    "references",
  ]

  const createdDirs = []
  for (const relativeDir of requiredDirs) {
    const absoluteDir = path.join(cwd, relativeDir)
    if (!fs.existsSync(absoluteDir)) {
      ensureDir(absoluteDir)
      createdDirs.push(relativeDir)
    }
  }

  if (createdDirs.length > 0) {
    log("✔", "Estructura minima creada (solo carpetas faltantes):")
    createdDirs.forEach(d => log("+", d))
  } else {
    log("·", "Estructura minima ya existente (sin cambios)")
  }

  // 4. npm install en access-vba-sync (por si hay nuevas dependencias)
  const syncDir = path.join(cwd, skillsDir, "access-vba-sync")
  if (fs.existsSync(path.join(syncDir, "package.json"))) {
    try {
      execSync("npm install --silent", { cwd: syncDir, stdio: "inherit" })
      log("✔", "Dependencias de access-vba-sync actualizadas")
    } catch (_) {
      log("✗", "Error en npm install — instala manualmente en " + skillsDir + "/access-vba-sync")
    }
  }

  // 5. Actualizar versión registrada
  fs.writeFileSync(versionFile, JSON.stringify({ version: pkgVersion }), "utf8")

  // Resumen
  console.log("\n  " + "=".repeat(40))
  console.log(`  ✅  Skills actualizadas a v${pkgVersion}`)
  console.log("  " + "=".repeat(40))
  console.log("\n  Archivos del proyecto NO modificados:")
  console.log("    · AGENTS.md")
  console.log("    · references/project_context.md")
  console.log("    · Contenido de docs/PRD/, docs/specs/, docs/plans/, docs/ERD/")
  console.log("    · src/, data/, .engram/")
  console.log("\n  Reinicia tu IDE para cargar las skills actualizadas.\n")
}
