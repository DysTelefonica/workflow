const { git } = require("../utils/git")
const { nextRelease } = require("../utils/version")

module.exports = function() {

  const version = nextRelease()

  console.log("Release:", version)

  git("checkout develop")
  git("pull")

  git("checkout main")
  git("pull")

  git("merge develop")

  git(`tag ${version}`)

  git("push origin main")
  git(`push origin ${version}`)

  console.log("Release created:", version)
}