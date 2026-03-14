const fs = require("fs")
const path = require("path")

function ensureDir(p) {
    if (!fs.existsSync(p)) {
        fs.mkdirSync(p, { recursive: true })
    }
}

module.exports = function () {

    const root = process.cwd()

    const dirs = [
        "docs",
        "docs/PRD",
        "docs/specs",
        "docs/specs/active",
        "docs/specs/completed",
        "src",
        "data",
        "rules",
        "skills",
        "scripts"
    ]

    dirs.forEach(d => ensureDir(path.join(root, d)))

    console.log("✔ Project structure created")

}