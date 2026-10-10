import type { Result } from "#result"
import { createResult, createResultError } from "#result"
import type { SharepointClient } from "./sharepointClientCreate.js"
import type { SharepointEntry } from "./sharepointEntrySchema.js"
import { sharepointErrorRead } from "./sharepointErrorRead.js"
import { sharepointItemRef } from "./sharepointItemRef.js"
import { sharepointPathFormat } from "./sharepointPathFormat.js"
import { sharepointPathParse } from "./sharepointPathParse.js"

const itemFields = ["name", "size", "file", "folder", "lastModifiedDateTime", "webUrl"] as const

/**
 * "" → sites (or libraries when a site is configured), "<site>" → libraries, "<site>/<library>[/folder]" → children.
 * A file path lists just that file.
 */
export async function sharepointList(client: SharepointClient, path: string = ""): Promise<Result<SharepointEntry[]>> {
  const op = "sharepointList"
  const site = client.config.site
  const segments = path.split("/").filter(Boolean)
  const librariesOf = (siteKey: string) =>
    client.libraries.filter((library) => library.siteKey.toLowerCase() === siteKey.toLowerCase())
  if (!site && segments.length === 0) {
    const sites = [...new Set(client.libraries.map((library) => library.siteKey))]
    return createResult(sites.map((key) => ({ path: `${key}/`, type: "site" as const })))
  }
  if ((site && segments.length === 0) || (!site && segments.length === 1)) {
    const siteKey = site ?? segments[0] ?? ""
    const libraries = librariesOf(siteKey)
    if (libraries.length === 0) return createResultError(op, `Unknown site "${siteKey}"`)
    return createResult(
      libraries.map((library) => ({
        path: sharepointPathFormat(library, "", { site, folder: true }),
        type: "library" as const,
        webUrl: library.driveUrl,
      })),
    )
  }
  const parsed = sharepointPathParse(path, client.libraries, { site })
  if (!parsed.success) return parsed
  const { library, itemPath } = parsed.data
  const label = sharepointPathFormat(library, itemPath, { site })
  try {
    const item = await sharepointItemRef(client.graph, library.driveId, itemPath).select(...itemFields)()
    if (item.file)
      return createResult([
        {
          path: label,
          type: "file",
          size: item.size ?? 0,
          modified: item.lastModifiedDateTime ?? undefined,
          webUrl: item.webUrl ?? undefined,
        },
      ])
    const entries: SharepointEntry[] = []
    const children = sharepointItemRef(client.graph, library.driveId, itemPath)
      .children.select(...itemFields)
      .top(200)
    for await (const page of children) {
      for (const child of page) {
        const childPath = [itemPath, child.name].filter(Boolean).join("/")
        entries.push({
          path: sharepointPathFormat(library, childPath, { site, folder: Boolean(child.folder) }),
          type: child.folder ? "folder" : "file",
          size: child.folder ? undefined : (child.size ?? 0),
          modified: child.lastModifiedDateTime ?? undefined,
          webUrl: child.webUrl ?? undefined,
        })
      }
    }
    return createResult(entries)
  } catch (error) {
    return createResultError(op, `${label}: ${sharepointErrorRead(error).message}`, String(error))
  }
}
