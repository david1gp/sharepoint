import type { Result } from "#result"
import { createResult, createResultError } from "#result"
import type { SharepointClient } from "./sharepointClientCreate.js"
import { sharepointErrorRead } from "./sharepointErrorRead.js"
import { sharepointItemPathEncode } from "./sharepointItemPathEncode.js"
import { sharepointPathFormat } from "./sharepointPathFormat.js"
import { sharepointPathParse } from "./sharepointPathParse.js"

/** Moves a file (or, with `recursive`, a folder with its contents) to the site's recycle bin. */
export async function sharepointDelete(
  client: SharepointClient,
  options: Readonly<{ path: string; recursive?: boolean }>,
): Promise<Result<string>> {
  const op = "sharepointDelete"
  const site = client.config.site
  const parsed = sharepointPathParse(options.path, client.libraries, { site })
  if (!parsed.success) return parsed
  const { library, itemPath } = parsed.data
  if (!itemPath) return createResultError(op, "Refusing to delete a whole library")
  const label = sharepointPathFormat(library, itemPath, { site })
  try {
    const item = client.graph.drives.getById(library.driveId).getItemByPath(sharepointItemPathEncode(itemPath))
    const info = await item.select("folder")()
    if (info.folder && !options.recursive)
      return createResultError(op, `${label} is a folder; use --recursive to delete it with its contents`)
    await item.delete()
    return createResult(label)
  } catch (error) {
    return createResultError(op, `${label}: ${sharepointErrorRead(error).message}`, String(error))
  }
}
