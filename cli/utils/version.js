const { git } = require("./git")

function nextRelease() {
  const tags = git("tag").split("\n")

  const year = new Date().getFullYear()

  const numbers = tags
    .filter(t => t.startsWith(year))
    .map(t => parseInt(t.split("-")[1]))

  const next = numbers.length ? Math.max(...numbers) + 1 : 1

  return `${year}-${String(next).padStart(3,"0")}`
}

module.exports = { nextRelease }