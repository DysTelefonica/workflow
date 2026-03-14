const { execSync } = require("child_process")

function git(cmd) {
  try {
    return execSync(`git ${cmd}`, { encoding: "utf8", stdio: "pipe" }).trim()
  } catch (err) {
    if (typeof err.status === "number") {
      const stderr = err.stderr ? String(err.stderr).trim() : err.message
      console.error(`Git error: ${stderr}`)
      process.exit(1)
    }
    throw new Error(`Git no disponible o no es un repositorio: ${err.message}`)
  }
}

module.exports = { git }