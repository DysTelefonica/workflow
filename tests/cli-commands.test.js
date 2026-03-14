const test = require("node:test")
const assert = require("node:assert/strict")
const fs = require("node:fs")
const os = require("node:os")
const path = require("node:path")
const { spawnSync } = require("node:child_process")

const repoRoot = path.resolve(__dirname, "..")
const cliPath = path.join(repoRoot, "cli", "workflow.js")

function mkTempDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), "dysflow-test-"))
}

function runCli(args, cwd, extraEnv = {}) {
  return spawnSync(process.execPath, [cliPath, ...args], {
    cwd,
    encoding: "utf8",
    env: { ...process.env, ...extraEnv }
  })
}

function writeFakeGitBin(tempProjectDir) {
  const fakeBinDir = path.join(tempProjectDir, "fake-bin")
  fs.mkdirSync(fakeBinDir, { recursive: true })

  const fakeGitCmd = path.join(fakeBinDir, "git.cmd")
  fs.writeFileSync(fakeGitCmd, "@echo off\r\nexit /b 0\r\n", "utf8")

  return fakeBinDir
}

test("spec: missing docs/specs/active returns clear error and exit code 1", (t) => {
  const tempDir = mkTempDir()
  t.after(() => fs.rmSync(tempDir, { recursive: true, force: true }))

  const result = runCli(["spec", "1"], tempDir)

  assert.equal(result.status, 1)
  assert.match(result.stdout, /docs\/specs\/active no existe\./)
})

test("plan: missing docs/plans/active returns clear error and exit code 1", (t) => {
  const tempDir = mkTempDir()
  t.after(() => fs.rmSync(tempDir, { recursive: true, force: true }))

  const result = runCli(["plan", "1"], tempDir)

  assert.equal(result.status, 1)
  assert.match(result.stdout, /docs\/plans\/active no existe\./)
})

test("spec: exact matching uses spec-1 and not spec-10", (t) => {
  const tempDir = mkTempDir()
  t.after(() => fs.rmSync(tempDir, { recursive: true, force: true }))
  const specsDir = path.join(tempDir, "docs", "specs", "active")
  fs.mkdirSync(specsDir, { recursive: true })

  fs.mkdirSync(path.join(specsDir, "spec-10-otra"))
  fs.mkdirSync(path.join(specsDir, "spec-1-correcta"))

  const fakeBin = writeFakeGitBin(tempDir)
  const result = runCli(["spec", "1"], tempDir, {
    PATH: `${fakeBin};${process.env.PATH || ""}`
  })

  assert.equal(result.status, 0)
  assert.match(result.stdout, /Creating branch: spec-1-correcta/)
})

test("plan: exact matching uses plan-1 and not plan-10", (t) => {
  const tempDir = mkTempDir()
  t.after(() => fs.rmSync(tempDir, { recursive: true, force: true }))
  const plansDir = path.join(tempDir, "docs", "plans", "active")
  fs.mkdirSync(plansDir, { recursive: true })

  fs.mkdirSync(path.join(plansDir, "plan-10-otro"))
  fs.mkdirSync(path.join(plansDir, "plan-1-correcto"))

  const fakeBin = writeFakeGitBin(tempDir)
  const result = runCli(["plan", "1"], tempDir, {
    PATH: `${fakeBin};${process.env.PATH || ""}`
  })

  assert.equal(result.status, 0)
  assert.match(result.stdout, /Creating branch: plan-1-correcto/)
})
