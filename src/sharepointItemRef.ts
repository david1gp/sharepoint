import type { SharepointGraph } from "./sharepointGraphCreate.js"
import { sharepointItemPathEncode } from "./sharepointItemPathEncode.js"

/** Graph drive item by path; an empty path is the library root folder. */
export function sharepointItemRef(graph: SharepointGraph, driveId: string, itemPath: string) {
  const drive = graph.drives.getById(driveId)
  return itemPath ? drive.getItemByPath(sharepointItemPathEncode(itemPath)) : drive.root
}
