const { execSync } = require("child_process")

function git(cmd) {
  return execSync(`git ${cmd}`, { encoding: "utf8" }).trim()
}

module.exports = { git }