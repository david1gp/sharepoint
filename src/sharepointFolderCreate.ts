import type { Result } from "#result"
import { createResult, createResultError } from "#result"
import type { SharepointClient } from "./sharepointClientCreate.js"
import { sharepointErrorRead } from "./sharepointErrorRead.js"
import { sharepointItemRef } from "./sharepointItemRef.js"
import { sharepointNotFoundIgnore } from "./sharepointNotFoundIgnore.js"
import { sharepointPathFormat } from "./sharepointPathFormat.js"
import { sharepointPathParse } from "./sharepointPathParse.js"

/** Creates the folder and any missing parents (like `mkdir -p`); existing folders are kept. */
export async function sharepointFolderCreate(client: SharepointClient, path: string): Promise<Result<string>> {
  const op = "sharepointFolderCreate"
  const site = client.config.site
  const parsed = sharepointPathParse(path, client.libraries, { site })
  if (!parsed.success) return parsed
  const { library, itemPath } = parsed.data
  if (!itemPath) return createResultError(op, "Give a folder path inside the library")
  const segments = itemPath.split("/")
  try {
    for (let index = 0; index < segments.length; index++) {
      const parent = segments.slice(0, index).join("/")
      const current = segments.slice(0, index + 1).join("/")
      const existing = await sharepointNotFoundIgnore(
        sharepointItemRef(client.graph, library.driveId, current).select("folder")(),
      )
      if (existing && !existing.folder)
        return createResultError(op, `${sharepointPathFormat(library, current, { site })} exists and is a file`)
      if (existing) continue
      await sharepointItemRef(client.graph, library.driveId, parent).children.addFolder({
        name: segments[index] ?? "",
        conflictBehavior: "fail",
      })
    }
    return createResult(sharepointPathFormat(library, itemPath, { site, folder: true }))
  } catch (error) {
    return createResultError(op, `${path}: ${sharepointErrorRead(error).message}`, String(error))
  }
}
