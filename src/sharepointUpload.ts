import type { Result } from "#result"
import { createResult, createResultError } from "#result"
import type { SharepointClient } from "./sharepointClientCreate.js"
import { sharepointErrorRead } from "./sharepointErrorRead.js"
import { sharepointItemPathEncode } from "./sharepointItemPathEncode.js"
import { sharepointItemRef } from "./sharepointItemRef.js"
import { sharepointNotFoundIgnore } from "./sharepointNotFoundIgnore.js"
import { sharepointPathFormat } from "./sharepointPathFormat.js"
import { sharepointPathParse } from "./sharepointPathParse.js"

type SharepointUploadOptions = Readonly<{
  /** Target file, or a folder ending in "/" (then `fileName` is appended). */
  path: string
  fileName: string
  content: Uint8Array
  /** Replace an existing file; without it an existing target is an error. */
  overwrite?: boolean
}>

// Graph's single PUT accepts up to 250 MB; larger files use an upload session in 320 KiB-multiple chunks.
const simpleUploadLimit = 250 * 1024 * 1024
const chunkSize = 10 * 320 * 1024

/** Uploads a file; missing parent folders are created by Graph. */
export async function sharepointUpload(
  client: SharepointClient,
  options: SharepointUploadOptions,
): Promise<Result<{ path: string; size: number; webUrl: string }>> {
  const op = "sharepointUpload"
  const site = client.config.site
  const parsed = sharepointPathParse(options.path, client.libraries, { site })
  if (!parsed.success) return parsed
  const { library, trailingSlash } = parsed.data
  const intoFolder = !parsed.data.itemPath || trailingSlash
  const itemPath = intoFolder
    ? [parsed.data.itemPath, options.fileName].filter(Boolean).join("/")
    : parsed.data.itemPath
  const label = sharepointPathFormat(library, itemPath, { site })
  try {
    const existing = await sharepointNotFoundIgnore(
      sharepointItemRef(client.graph, library.driveId, itemPath).select("file", "folder")(),
    )
    if (existing?.folder) return createResultError(op, `${label} is a folder`)
    // PnP's upload() cannot pass Graph's conflictBehavior, so refuse existing targets explicitly.
    if (existing && !options.overwrite)
      return createResultError(op, `${label} already exists; use --overwrite to replace it`)
    const total = options.content.byteLength
    if (total <= simpleUploadLimit) {
      const item = await sharepointItemRef(client.graph, library.driveId, "").upload({
        filePathName: sharepointItemPathEncode(itemPath),
        content: options.content,
      })
      return createResult({ path: label, size: item.size ?? total, webUrl: item.webUrl ?? "" })
    }
    const { resumableUpload } = await client.graph.drives
      .getById(library.driveId)
      .getItemByPath(sharepointItemPathEncode(itemPath))
      .createUploadSession({})
    let last: Record<string, unknown> = {}
    for (let start = 0; start < total; start += chunkSize) {
      const end = Math.min(start + chunkSize, total)
      const chunk = options.content.slice(start, end)
      // The final chunk's response is the created driveItem (size, webUrl), not the session typed by PnP.
      last = (await resumableUpload.upload(end - start, chunk, `bytes ${start}-${end - 1}/${total}`)) as Record<
        string,
        unknown
      >
    }
    return createResult({
      path: label,
      size: typeof last.size === "number" ? last.size : total,
      webUrl: typeof last.webUrl === "string" ? last.webUrl : "",
    })
  } catch (error) {
    return createResultError(op, `${label}: ${sharepointErrorRead(error).message}`, String(error))
  }
}
