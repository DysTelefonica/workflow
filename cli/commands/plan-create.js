const { git } = require("../utils/git")
const fs = require("fs")

module.exports = function(planNumber) {
  const plansDir = "docs/plans/active"
  if (!fs.existsSync(plansDir)) {
    console.log("Error: El directorio docs/plans/active no existe.")
    console.log("Ejecuta 'dysflow init access' primero.")
    process.exit(1)
  }

  const folders = fs.readdirSync(plansDir)
  const exactPlanPrefix = new RegExp(`^plan-${String(planNumber)}(?:-|$)`)

  const plan = folders.find(f => exactPlanPrefix.test(f))

  if (!plan) {
    console.log("Plan not found in docs/plans/active/")
    process.exit(1)
  }

  const branch = plan

  console.log("Creating branch:", branch)

  git("checkout develop")
  git("pull")

  git(`checkout -b ${branch}`)

  git(`push -u origin ${branch}`)
}
