#!/usr/bin/env bun

import { run } from "@stricli/core"
import { sharepointApplication } from "./sharepointApplication.js"

// Stricli reports usage errors with negative codes; map them to conventional exit codes (2 usage, 1 failure).
const applicationProcess = {
  env: process.env,
  stderr: process.stderr,
  stdout: process.stdout,
  get exitCode() {
    return process.exitCode
  },
  set exitCode(value: number | string | null | undefined) {
    const numericValue = typeof value === "number" ? value : Number(value)
    if (numericValue === -4 || numericValue === -5) {
      process.exitCode = 2
      return
    }
    if (numericValue < 0) {
      process.exitCode = 1
      return
    }
    process.exitCode = value
  },
}

await run(sharepointApplication, process.argv.slice(2), { process: applicationProcess })
