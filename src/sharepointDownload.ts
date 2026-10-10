import type { Result } from "#result"
import { createResult, createResultError } from "#result"
import type { SharepointClient } from "./sharepointClientCreate.js"
import { sharepointErrorRead } from "./sharepointErrorRead.js"
import { sharepointItemPathEncode } from "./sharepointItemPathEncode.js"
import { sharepointPathParse } from "./sharepointPathParse.js"

/** File contents of `<site>/<library>/<file>`. */
export async function sharepointDownload(
  client: SharepointClient,
  path: string,
): Promise<Result<{ name: string; content: Uint8Array }>> {
  const op = "sharepointDownload"
  const parsed = sharepointPathParse(path, client.libraries, { site: client.config.site })
  if (!parsed.success) return parsed
  const { library, itemPath } = parsed.data
  if (!itemPath) return createResultError(op, `"${path}" is a library, not a file`)
  try {
    const item = () => client.graph.drives.getById(library.driveId).getItemByPath(sharepointItemPathEncode(itemPath))
    const info = await item().select("file")()
    if (!info.file) return createResultError(op, `"${path}" is a folder, not a file`)
    // Fresh reference: getContent() needs the unselected item's @microsoft.graph.downloadUrl.
    const blob = await item().getContent()
    return createResult({
      name: itemPath.split("/").at(-1) ?? itemPath,
      content: new Uint8Array(await blob.arrayBuffer()),
    })
  } catch (error) {
    return createResultError(op, `${path}: ${sharepointErrorRead(error).message}`, String(error))
  }
}
