import * as v from "valibot"

/** One search hit; `path` is a `sharepoint get`-able path when the file lies in a known library, else the web URL. */
export const sharepointSearchHitSchema = v.object({
  path: v.string(),
  webUrl: v.string(),
  size: v.optional(v.number()),
  modified: v.optional(v.string()),
  summary: v.string(),
})

export type SharepointSearchHit = v.InferOutput<typeof sharepointSearchHitSchema>
