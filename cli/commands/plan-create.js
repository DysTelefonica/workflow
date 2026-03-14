const { git } = require("../utils/git")

module.exports = function(planNumber) {

  const folders = require("fs").readdirSync("docs/plans/active")

  const plan = folders.find(f => f.startsWith(`plan-${planNumber}`))

  if (!plan) {
    console.log("Plan not found in docs/plans/active/")
    process.exit()
  }

  const branch = plan

  console.log("Creating branch:", branch)

  git("checkout develop")
  git("pull")

  git(`checkout -b ${branch}`)

  git(`push -u origin ${branch}`)
}
