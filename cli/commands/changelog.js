const { git } = require("../utils/git")

module.exports = function(from){

  const log = git(`log ${from}..HEAD --pretty=format:"%s"`)

  const lines = log.split("\n")

  console.log("Changelog\n")

  lines
    .filter(l => l.includes("spec-"))
    .forEach(l => console.log("-", l))
}