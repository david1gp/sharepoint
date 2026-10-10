import * as v from "valibot"

/** One line of `sharepoint ls`: a site, a library, a folder or a file. */
export const sharepointEntrySchema = v.object({
  path: v.string(),
  type: v.picklist(["site", "library", "folder", "file"]),
  size: v.optional(v.number()),
  modified: v.optional(v.string()),
  webUrl: v.optional(v.string()),
})

export type SharepointEntry = v.InferOutput<typeof sharepointEntrySchema>
