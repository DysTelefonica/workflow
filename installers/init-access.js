#!/usr/bin/env node
"use strict"

const fs   = require("fs")
const path = require("path")
const { execSync } = require("child_process")
const readline = require("readline")

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function rl() {
  return readline.createInterface({ input: process.stdin, output: process.stdout })
}

function ask(iface, question) {
  return new Promise(resolve => iface.question(question, answer => resolve(answer.trim())))
}

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
// Access DB detection
// ---------------------------------------------------------------------------

function scanAccess(cwd) {
  const exts    = [".accdb", ".accde", ".mdb"]
  const dataRx  = /_[Dd]atos?/
  const files   = fs.readdirSync(cwd).filter(f => exts.includes(path.extname(f).toLowerCase()))
  const front   = files.filter(f => !dataRx.test(path.basename(f, path.extname(f))))
  const back    = files.filter(f =>  dataRx.test(path.basename(f, path.extname(f))))
  return { front, back }
}

async function resolveAccessFiles(iface, cwd) {
  let { front, back } = scanAccess(cwd)

  // Si no hay ningún archivo Access → pedir que los copien
  if (front.length === 0 && back.length === 0) {
    console.log("\n  ⚠  No se encontraron archivos Access en esta carpeta.\n")
    console.log("     Copia aquí antes de continuar:")
    console.log("       → El frontend:  PROYECTO.accdb        (sin _datos en el nombre)")
    console.log("       → El backend:   PROYECTO_Datos.accdb   (con _datos en el nombre)\n")
    await ask(iface, "     Pulsa Enter cuando estés listo...")
    ;({ front, back } = scanAccess(cwd))
    if (front.length === 0 && back.length === 0) {
      log("·", "Ningún archivo Access encontrado — se omitirá VBA export y ERD.")
      return { frontend: null, backends: [] }
    }
  }

  // Elegir frontend si hay varios
  let frontend = null
  if (front.length === 1) {
    frontend = front[0]
    log("✔", `Frontend detectado: ${frontend}`)
  } else if (front.length > 1) {
    console.log("\n  Se encontraron varios archivos frontend:")
    front.forEach((f, i) => console.log(`    ${i + 1}) ${f}`))
    const choice = await ask(iface, "  ¿Cuál es el principal? (número): ")
    const idx = parseInt(choice, 10) - 1
    frontend = front[idx] || front[0]
    log("✔", `Frontend seleccionado: ${frontend}`)
  }

  if (back.length === 0) {
    log("·", "No se encontró *_Datos.accdb — ERD se omitirá.")
    log("·", `Puedes generarlo manualmente: node <SKILLS_DIR>/access-vba-sync/cli.js generate-erd --backend <ruta>`)
  } else {
    back.forEach(b => log("✔", `Backend detectado: ${b}`))
  }

  return { frontend, backends: back }
}

// ---------------------------------------------------------------------------
// Skills dir detection
// ---------------------------------------------------------------------------

async function resolveSkillsDir(iface, cwd) {
  // Auto-detect en proyecto existente
  if (fs.existsSync(path.join(cwd, ".agents"))) {
    log("✔", "Detectado: Trae nuevo (.agents/)")
    return { skillsDir: ".agents", rulesDir: ".agents/rules", detected: true }
  }
  if (fs.existsSync(path.join(cwd, ".trae", "skills"))) {
    console.log("\n  ⚠  Detectado formato Trae antiguo (.trae/skills).")
    const ans = await ask(iface, "  ¿Migrar a .agents/ (nuevo)? S/N: ")
    if (/^s/i.test(ans)) {
      log("→", "Se migrará a .agents/ (el directorio .trae/ no se borrará)")
      return { skillsDir: ".agents", rulesDir: ".agents/rules", detected: true, migrating: true, oldDir: ".trae/skills" }
    }
    return { skillsDir: ".trae/skills", rulesDir: ".trae/rules", detected: true }
  }
  if (fs.existsSync(path.join(cwd, ".agent", "skills"))) {
    log("✔", "Detectado: formato .agent/skills/ (template antiguo)")
    return { skillsDir: ".agent/skills", rulesDir: ".agent/rules", detected: true }
  }
  if (fs.existsSync(path.join(cwd, "skills"))) {
    log("✔", "Detectado: Standard (skills/)")
    return { skillsDir: "skills", rulesDir: "rules", detected: true }
  }

  // Proyecto nuevo — preguntar
  console.log("\n  ¿Qué IDE estás usando?")
  console.log("    1) Trae   → .agents/")
  console.log("    2) Claude Code / Standard → skills/")
  const choice = await ask(iface, "  IDE (1/2): ")
  if (choice === "1") {
    return { skillsDir: ".agents", rulesDir: ".agents/rules", detected: false }
  }
  return { skillsDir: "skills", rulesDir: "rules", detected: false }
}

// ---------------------------------------------------------------------------
// Template rendering
// ---------------------------------------------------------------------------

function renderTemplate(templatePath, vars) {
  if (!fs.existsSync(templatePath)) return null
  let content = fs.readFileSync(templatePath, "utf8")
  for (const [key, value] of Object.entries(vars)) {
    content = content.replaceAll(`{{${key}}}`, value)
  }
  return content
}

// ---------------------------------------------------------------------------
// Main installer
// ---------------------------------------------------------------------------

module.exports = async function initAccess(opts = {}) {
  const cwd      = process.cwd()
  const pkgRoot  = path.join(__dirname, "..")   // raíz del paquete workflow
  const iface    = rl()

  console.log("\n\n  🛠   WORKFLOW — INIT ACCESS\n" + "  " + "=".repeat(40))

  try {
    // 1. Detectar Skills dir
    const { skillsDir, rulesDir, migrating, oldDir } = await resolveSkillsDir(iface, cwd)

    // 2. Nombre del proyecto (si proyecto nuevo o si no hay AGENTS.md)
    let projectName = ""
    let projectDomain = ""
    let projectStage = ""
    const agentsExists = fs.existsSync(path.join(cwd, "AGENTS.md"))

    if (!agentsExists) {
      console.log("")
      projectName   = await ask(iface, "  Nombre del proyecto (ej: CONDOR): ")
      projectDomain = await ask(iface, "  Dominio (ej: Gestión de mantenimiento): ")
      projectStage  = await ask(iface, "  Fase actual (ej: Autodescubrimiento): ")
    }

    // 3. Detectar archivos Access
    console.log("\n  --- Detectando archivos Access ---")
    const { frontend, backends } = await resolveAccessFiles(iface, cwd)

    // 4. Crear estructura de carpetas
    console.log("\n  --- Creando estructura ---")
    const dirs = [
      "docs/PRD",
      "docs/specs/active",
      "docs/specs/completed",
      "docs/templates",
      "docs/ERD",
      "src/modules",
      "src/classes",
      "src/forms",
      "data",
      ".engram",
      skillsDir,
      rulesDir,
    ]
    dirs.forEach(d => {
      ensureDir(path.join(cwd, d))
      log("+", d)
    })

    // 5. Copiar framework desde el paquete
    console.log("\n  --- Copiando framework ---")

    // Skills
    const srcSkills = path.join(pkgRoot, "skills")
    const destSkills = path.join(cwd, skillsDir)
    copyDir(srcSkills, destSkills)
    log("✔", `skills/ → ${skillsDir}`)

    // Rules
    const srcRules = path.join(pkgRoot, "rules")
    const destRules = path.join(cwd, rulesDir)
    copyDir(srcRules, destRules)
    log("✔", `rules/ → ${rulesDir}`)

    // Templates
    const srcTemplates = path.join(pkgRoot, "templates")
    const destTemplates = path.join(cwd, "docs/templates")
    copyDir(srcTemplates, destTemplates)
    log("✔", `templates/ → docs/templates/`)

    // 6. AGENTS.md (solo si no existe)
    if (!agentsExists) {
      console.log("\n  --- Generando AGENTS.md ---")
      const templateVars = {
        PROJECT_NAME:   projectName,
        PROJECT_DOMAIN: projectDomain,
        PROJECT_STAGE:  projectStage,
        SKILLS_DIR:     skillsDir,
        RULES_DIR:      rulesDir,
      }
      const agentsContent = renderTemplate(path.join(pkgRoot, "templates", "AGENTS_template.md"), templateVars)
      if (agentsContent) {
        fs.writeFileSync(path.join(cwd, "AGENTS.md"), agentsContent, "utf8")
        log("✔", `AGENTS.md generado para "${projectName}"`)
      } else {
        log("✗", "AGENTS_template.md no encontrado en el paquete")
      }

      // project_context.md desde prd-writer
      const ctxTemplate = path.join(cwd, skillsDir, "prd-writer", "references", "project_context_template.md")
      const ctxDest     = path.join(cwd, "references", "project_context.md")
      ensureDir(path.join(cwd, "references"))
      if (!fs.existsSync(ctxDest) && fs.existsSync(ctxTemplate)) {
        let ctx = fs.readFileSync(ctxTemplate, "utf8")
        ctx = ctx.replaceAll("{NOMBRE_PROYECTO}", projectName)
        fs.writeFileSync(ctxDest, ctx, "utf8")
        log("✔", "references/project_context.md creado")
      }
    } else {
      log("→", "AGENTS.md ya existe — sin cambios")
    }

    // 7. npm install en access-vba-sync
    console.log("\n  --- npm install (access-vba-sync) ---")
    const syncDir = path.join(cwd, skillsDir, "access-vba-sync")
    if (fs.existsSync(path.join(syncDir, "package.json"))) {
      try {
        execSync("npm install --silent", { cwd: syncDir, stdio: "inherit" })
        log("✔", "Dependencias instaladas")
      } catch (e) {
        log("✗", "Error en npm install — instala manualmente en " + skillsDir + "/access-vba-sync")
      }
    } else {
      log("·", "package.json no encontrado en access-vba-sync — omitido")
    }

    // 8. Export VBA inicial
    const cliJs = path.join(syncDir, "cli.js")
    if (frontend && fs.existsSync(cliJs)) {
      console.log("\n  --- Export VBA inicial ---")
      try {
        execSync(`node "${cliJs}" start --access "${path.join(cwd, frontend)}"`, { cwd, stdio: "inherit" })
        log("✔", "Módulos exportados a src/")
      } catch (e) {
        log("✗", "Error en VBA export — ejecuta manualmente: node " + skillsDir + "/access-vba-sync/cli.js start")
      }
    }

    // 9. ERD por cada backend
    if (backends.length > 0 && fs.existsSync(cliJs)) {
      console.log("\n  --- Generando ERD ---")
      const erdDir = path.join(cwd, "docs", "ERD")
      ensureDir(erdDir)
      for (const b of backends) {
        const erdName = path.basename(b, path.extname(b)) + ".md"
        try {
          execSync(
            `node "${cliJs}" generate-erd --backend "${path.join(cwd, b)}" --erd_path "${erdDir}"`,
            { cwd, stdio: "inherit" }
          )
          log("✔", `ERD generado: docs/ERD/${erdName}`)
        } catch (e) {
          log("✗", `Error generando ERD para ${b}`)
        }
      }
    }

    // 10. Nota migración
    if (migrating && oldDir) {
      console.log("\n  ⚠  El directorio antiguo " + oldDir + " se ha mantenido.")
      console.log("     Puedes eliminarlo cuando compruebes que todo funciona.")
    }

    // 11. Escribir version del framework instalado
    const pkgVersion = require("../package.json").version
    fs.writeFileSync(
      path.join(cwd, skillsDir, ".dysflow"),
      JSON.stringify({ version: pkgVersion }),
      "utf8"
    )

    // Resumen
    console.log("\n  " + "=".repeat(40))
    console.log("  ✅  LISTO — " + (projectName || path.basename(cwd)))
    console.log("  " + "=".repeat(40))
    console.log(`  Skills  : ${skillsDir}`)
    console.log(`  Rules   : ${rulesDir}`)
    if (frontend) console.log(`  Frontend: ${frontend}`)
    backends.forEach(b => console.log(`  Backend : ${b}`))
    console.log("\n  Próximos pasos:")
    console.log("  1. Revisa AGENTS.md")
    console.log("  2. Reinicia tu IDE para cargar las skills")
    console.log("  3. Verifica src/ y docs/ERD/ generados\n")

  } finally {
    iface.close()
  }
}
