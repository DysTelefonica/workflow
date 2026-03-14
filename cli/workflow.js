#!/usr/bin/env node

const { Command } = require("commander")

const specCreate = require("./commands/spec-create")
const release = require("./commands/release")
const hotfix = require("./commands/hotfix")
const changelog = require("./commands/changelog")
const nextRelease = require("./commands/next-release")

const program = new Command()

program.name("workflow")

program
  .command("spec <number>")
  .action(specCreate)

program
  .command("release")
  .action(release)

program
  .command("hotfix <name>")
  .action(hotfix)

program
  .command("changelog <from>")
  .action(changelog)

program
  .command("next-release")
  .action(nextRelease)

program.parse(process.argv)

const initProject = require("./commands/init-project")

program
  .command("init")
  .description("Initialize workflow project")
  .action(initProject)