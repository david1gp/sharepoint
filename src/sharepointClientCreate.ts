import type { Result } from "#result"
import { createResult } from "#result"
import type { SharepointConfig } from "./sharepointConfigSchema.js"
import { type SharepointGraph, sharepointGraphCreate } from "./sharepointGraphCreate.js"
import { sharepointLibrariesList } from "./sharepointLibrariesList.js"
import type { SharepointLibrary } from "./sharepointLibrarySchema.js"

export type SharepointClient = Readonly<{
  config: SharepointConfig
  graph: SharepointGraph
  libraries: readonly SharepointLibrary[]
}>

/** Signs in and discovers the document libraries once; every file operation takes the returned client. */
export async function sharepointClientCreate(
  config: SharepointConfig,
  options: Readonly<{ graph?: SharepointGraph; warn?: (message: string) => void }> = {},
): Promise<Result<SharepointClient>> {
  const graph = options.graph ?? sharepointGraphCreate(config)
  const libraries = await sharepointLibrariesList(graph, {
    hostname: config.hostname,
    site: config.site,
    warn: options.warn,
  })
  if (!libraries.success) return libraries
  return createResult({ config, graph, libraries: libraries.data })
}
