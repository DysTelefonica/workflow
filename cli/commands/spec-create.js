const { git } = require("../utils/git")
const fs = require("fs")

module.exports = function(specNumber) {
  const specsDir = "docs/specs/active"
  if (!fs.existsSync(specsDir)) {
    console.log("Error: El directorio docs/specs/active no existe.")
    console.log("Ejecuta 'dysflow init access' primero.")
    process.exit(1)
  }

  const folders = fs.readdirSync(specsDir)
  const exactSpecPrefix = new RegExp(`^spec-${String(specNumber)}(?:-|$)`)

  const spec = folders.find(f => exactSpecPrefix.test(f))

  if(!spec){
    console.log("Spec not found in docs/specs/active/")
    process.exit(1)
  }

  const branch = spec

  console.log("Creating branch:", branch)

  git("checkout develop")
  git("pull")

  git(`checkout -b ${branch}`)

  git(`push -u origin ${branch}`)
}
