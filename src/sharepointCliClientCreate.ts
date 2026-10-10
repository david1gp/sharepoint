import type { CommandContext } from "@stricli/core"
import * as v from "valibot"
import type { Result } from "#result"
import { createResultError } from "#result"
import { type SharepointClient, sharepointClientCreate } from "./sharepointClientCreate.js"
import { sharepointConfigLoad } from "./sharepointConfigLoad.js"
import { sharepointGraphCreate } from "./sharepointGraphCreate.js"

// Unauthenticated local Graph fake for CLI tests; refusing remote hosts keeps it from ever sending real data elsewhere.
const graphUrlSchema = v.pipe(v.string(), v.url(), v.regex(/^http:\/\/(localhost|127\.0\.0\.1):\d+\//))

/** Config (--config / SHAREPOINT_CONFIG / ~/.config/sharepoint/config.json / SHAREPOINT_*) → signed-in client. */
export async function sharepointCliClientCreate(
  context: CommandContext,
  flags: Readonly<{ config?: string }>,
): Promise<Result<SharepointClient>> {
  const op = "sharepointCliClientCreate"
  const config = await sharepointConfigLoad({ configPath: flags.config, environment: process.env })
  if (!config.success) return config
  const graphUrl = process.env.SHAREPOINT_GRAPH_URL
  const parsedUrl = graphUrl ? v.safeParse(graphUrlSchema, graphUrl) : undefined
  if (parsedUrl && !parsedUrl.success)
    return createResultError(op, "SHAREPOINT_GRAPH_URL must be a local http://localhost:<port>/ test URL")
  return sharepointClientCreate(config.data, {
    graph: parsedUrl?.success ? sharepointGraphCreate(config.data, { baseUrl: parsedUrl.output }) : undefined,
    warn: (message) => context.process.stderr.write(`warning: ${message}\n`),
  })
}
