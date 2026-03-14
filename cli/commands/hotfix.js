const { git } = require("../utils/git")

module.exports = function(name){

  const branch = `hotfix-${name}`

  git("checkout main")
  git("pull")

  git(`checkout -b ${branch}`)

  git(`push -u origin ${branch}`)

  console.log("Hotfix branch:", branch)
}