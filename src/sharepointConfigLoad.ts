import { readFile as readFileDefault } from "node:fs/promises"
import * as v from "valibot"
import type { Result } from "#result"
import { createResult, createResultError } from "#result"
import { sharepointConfigPathResolve } from "./sharepointConfigPathResolve.js"
import { type SharepointConfig, sharepointConfigSchema } from "./sharepointConfigSchema.js"
import { sharepointEnvironmentConfigRead } from "./sharepointEnvironmentConfigRead.js"

type SharepointConfigLoadOptions = Readonly<{
  /** Explicit config file (CLI --config). */
  configPath?: string
  environment?: Readonly<Record<string, string | undefined>>
  home?: string
  readFile?: (path: string, encoding: "utf8") => Promise<string>
}>

const sharepointConfigFileSchema = v.partial(sharepointConfigSchema)

/**
 * Merges the JSON config file with `SHAREPOINT_*` variables (variables win per field) and validates the result.
 * The default file may be missing when variables are complete; an explicit file (--config / SHAREPOINT_CONFIG) must exist.
 */
export async function sharepointConfigLoad(
  options: SharepointConfigLoadOptions = {},
): Promise<Result<SharepointConfig>> {
  const op = "sharepointConfigLoad"
  const environment = options.environment ?? process.env
  const readFile = options.readFile ?? readFileDefault
  const { path, explicit } = sharepointConfigPathResolve({ ...options, environment })
  let fileConfig: v.InferOutput<typeof sharepointConfigFileSchema> = {}
  let text: string | undefined
  try {
    text = await readFile(path, "utf8")
  } catch (error) {
    const missing = typeof error === "object" && error !== null && "code" in error && error.code === "ENOENT"
    if (explicit || !missing) return createResultError(op, `Cannot read config file ${path}: ${String(error)}`)
  }
  if (text !== undefined) {
    let json: unknown
    try {
      json = JSON.parse(text)
    } catch (error) {
      return createResultError(op, `${path} is not valid JSON: ${String(error)}`)
    }
    const parsedFile = v.safeParse(sharepointConfigFileSchema, json)
    if (!parsedFile.success) return createResultError(op, `${path}: ${v.summarize(parsedFile.issues)}`)
    fileConfig = parsedFile.output
  }
  const merged = { ...fileConfig, ...sharepointEnvironmentConfigRead(environment) }
  if (!merged.tenantId && !merged.clientId && !merged.clientSecret)
    return createResultError(op, `SharePoint is not configured: create ${path} or set SHAREPOINT_* variables`)
  const parsed = v.safeParse(sharepointConfigSchema, merged)
  if (!parsed.success) return createResultError(op, v.summarize(parsed.issues))
  return createResult(parsed.output)
}
