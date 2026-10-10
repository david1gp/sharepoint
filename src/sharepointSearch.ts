import type { Result } from "#result"
import { createResult, createResultError } from "#result"
import type { SharepointClient } from "./sharepointClientCreate.js"
import { sharepointErrorRead } from "./sharepointErrorRead.js"
import { sharepointPathFormat } from "./sharepointPathFormat.js"
import type { SharepointSearchHit } from "./sharepointSearchHitSchema.js"
import { sharepointSearchRegionRead } from "./sharepointSearchRegionRead.js"
import { sharepointUrlDecode } from "./sharepointUrlDecode.js"

type SearchRequest = Parameters<SharepointClient["graph"]["query"]>[0]
type SharepointSearchResult = { total: number; hits: SharepointSearchHit[]; region: string }

// Without a region Graph only says "Region is required"; an unknown one makes it list the valid regions.
const regionProbe = "XX"

/**
 * Full-text Microsoft Search over files (KQL: words, "phrases", filetype:pdf, path:"https://…").
 * The region is taken from the config or discovered on first use. A configured site restricts hits to that site.
 */
export async function sharepointSearch(
  client: SharepointClient,
  options: Readonly<{ query: string; limit?: number }>,
): Promise<Result<SharepointSearchResult>> {
  const op = "sharepointSearch"
  const site = client.config.site
  const siteUrl = site ? client.libraries[0]?.siteUrl : undefined
  const queryString = siteUrl ? `${options.query} path:"${siteUrl}"` : options.query
  const request = (region: string): SearchRequest => ({
    entityTypes: ["driveItem"],
    query: { queryString },
    region,
    from: 0,
    size: options.limit ?? 25,
  })
  let region = client.config.searchRegion ?? regionProbe
  try {
    const response = await client.graph.query(request(region)).catch(async (error: unknown) => {
      const discovered =
        region === regionProbe ? sharepointSearchRegionRead(sharepointErrorRead(error).message) : undefined
      if (!discovered) throw error
      region = discovered
      return client.graph.query(request(region))
    })
    const container = response[0]?.hitsContainers?.[0]
    const hits = (container?.hits ?? []).map((hit): SharepointSearchHit => {
      const resource = (hit.resource ?? {}) as {
        name?: string
        webUrl?: string
        size?: number
        lastModifiedDateTime?: string
        parentReference?: { driveId?: string }
      }
      const webUrl = resource.webUrl ?? ""
      const library = client.libraries.find((candidate) => candidate.driveId === resource.parentReference?.driveId)
      const libraryUrl = library ? sharepointUrlDecode(library.driveUrl) : ""
      const decodedUrl = sharepointUrlDecode(webUrl)
      const itemPath =
        library && decodedUrl.startsWith(`${libraryUrl}/`) ? decodedUrl.slice(libraryUrl.length + 1) : undefined
      return {
        path: library && itemPath ? sharepointPathFormat(library, itemPath, { site }) : webUrl,
        webUrl,
        size: resource.size,
        modified: resource.lastModifiedDateTime,
        summary: (hit.summary ?? "")
          .replace(/<\/?c0>/g, "")
          .replace(/<ddd\/>/g, " … ")
          .trim(),
      }
    })
    return createResult({ total: container?.total ?? hits.length, hits, region })
  } catch (error) {
    return createResultError(op, sharepointErrorRead(error).message, String(error))
  }
}
