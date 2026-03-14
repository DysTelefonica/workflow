#!/usr/bin/env node

const { Command } = require("commander")

const program = new Command()
program.name("dysflow")

program
  .command("spec <number>")
  .action(require("./commands/spec-create"))

program
  .command("release")
  .action(require("./commands/release"))

program
  .command("hotfix <name>")
  .action(require("./commands/hotfix"))

program
  .command("changelog <from>")
  .action(require("./commands/changelog"))

program
  .command("next-release")
  .action(require("./commands/next-release"))

program
  .command("init <type>")
  .description("Initialize workflow project. Types: access")
  .option("--migrate", "Migrate existing project to new skills location")
  .action((type, opts) => {
    if (type === "access") {
      require("../installers/init-access")(opts)
    } else {
      console.error(`Unknown project type: ${type}. Use: access`)
      process.exit(1)
    }
  })

program
  .command("update")
  .description("Actualiza skills y rules a la ultima version sin tocar el proyecto")
  .action(require("../installers/update-access"))

program.parse(process.argv)
