const { git } = require("../utils/git")

module.exports = function(specNumber) {

  const folders = require("fs").readdirSync("docs/specs/active")

  const spec = folders.find(f => f.startsWith(`spec-${specNumber}`))

  if(!spec){
    console.log("Spec not found")
    process.exit()
  }

  const branch = spec

  console.log("Creating branch:", branch)

  git("checkout develop")
  git("pull")

  git(`checkout -b ${branch}`)

  git(`push -u origin ${branch}`)
}