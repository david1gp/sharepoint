import { readFile, writeFile } from "node:fs/promises"
import { basename } from "node:path"
import type { CommandContext } from "@stricli/core"
import { sharepointCliClientCreate } from "./sharepointCliClientCreate.js"
import { sharepointCliError } from "./sharepointCliError.js"
import { sharepointDelete } from "./sharepointDelete.js"
import { sharepointDownload } from "./sharepointDownload.js"
import { sharepointFolderCreate } from "./sharepointFolderCreate.js"
import { sharepointList } from "./sharepointList.js"
import { sharepointSearch } from "./sharepointSearch.js"
import { sharepointUpload } from "./sharepointUpload.js"

type ConfigFlags = Readonly<{ config?: string }>
type OutputFlags = ConfigFlags & Readonly<{ json: boolean }>
type CommandOutcome = Promise<Error | undefined>

const json = (value: unknown) => `${JSON.stringify(value, null, 2)}\n`

/** Command handlers; `this` is the Stricli context (stdout/stderr), so tests can capture output. */
export const sharepointCommands = {
  async ls(this: CommandContext, flags: OutputFlags & Readonly<{ long: boolean }>, path?: string): CommandOutcome {
    const client = await sharepointCliClientCreate(this, flags)
    if (!client.success) return sharepointCliError(client)
    const result = await sharepointList(client.data, path ?? "")
    if (!result.success) return sharepointCliError(result)
    if (flags.json) {
      this.process.stdout.write(json(result.data))
      return undefined
    }
    for (const entry of result.data) {
      const line = flags.long
        ? `${String(entry.size ?? "-").padStart(12)}  ${(entry.modified ?? "").padEnd(20)}  ${entry.path}`
        : entry.path
      this.process.stdout.write(`${line}\n`)
    }
    return undefined
  },

  async get(this: CommandContext, flags: ConfigFlags, path: string, destination?: string): CommandOutcome {
    const client = await sharepointCliClientCreate(this, flags)
    if (!client.success) return sharepointCliError(client)
    const result = await sharepointDownload(client.data, path)
    if (!result.success) return sharepointCliError(result)
    if (destination === "-") {
      // Stricli's stdout only takes strings; binary file contents must go to the real stream.
      process.stdout.write(result.data.content)
      return undefined
    }
    const target = destination ?? result.data.name
    await writeFile(target, result.data.content)
    this.process.stderr.write(`${path} → ${target} (${result.data.content.byteLength} bytes)\n`)
    return undefined
  },

  async put(
    this: CommandContext,
    flags: ConfigFlags & Readonly<{ overwrite: boolean }>,
    file: string,
    path: string,
  ): CommandOutcome {
    const content = await readFile(file).catch(() => undefined)
    if (!content) return new Error(`sharepoint: cannot read local file: ${file}`)
    const client = await sharepointCliClientCreate(this, flags)
    if (!client.success) return sharepointCliError(client)
    const result = await sharepointUpload(client.data, {
      path,
      fileName: basename(file),
      content: new Uint8Array(content),
      overwrite: flags.overwrite,
    })
    if (!result.success) return sharepointCliError(result)
    this.process.stdout.write(`${file} → ${result.data.path} (${result.data.size} bytes)\n${result.data.webUrl}\n`)
    return undefined
  },

  async mkdir(this: CommandContext, flags: ConfigFlags, path: string): CommandOutcome {
    const client = await sharepointCliClientCreate(this, flags)
    if (!client.success) return sharepointCliError(client)
    const result = await sharepointFolderCreate(client.data, path)
    if (!result.success) return sharepointCliError(result)
    this.process.stdout.write(`${result.data}\n`)
    return undefined
  },

  async rm(this: CommandContext, flags: ConfigFlags & Readonly<{ recursive: boolean }>, path: string): CommandOutcome {
    const client = await sharepointCliClientCreate(this, flags)
    if (!client.success) return sharepointCliError(client)
    const result = await sharepointDelete(client.data, { path, recursive: flags.recursive })
    if (!result.success) return sharepointCliError(result)
    this.process.stdout.write(`deleted (moved to recycle bin): ${result.data}\n`)
    return undefined
  },

  async search(this: CommandContext, flags: OutputFlags & Readonly<{ limit: number }>, query: string): CommandOutcome {
    const client = await sharepointCliClientCreate(this, flags)
    if (!client.success) return sharepointCliError(client)
    const result = await sharepointSearch(client.data, { query, limit: flags.limit })
    if (!result.success) return sharepointCliError(result)
    if (flags.json) {
      this.process.stdout.write(json(result.data))
      return undefined
    }
    this.process.stderr.write(`${result.data.total} hit(s)\n`)
    for (const hit of result.data.hits) this.process.stdout.write(`${hit.path}\n    ${hit.summary}\n`)
    return undefined
  },
}
