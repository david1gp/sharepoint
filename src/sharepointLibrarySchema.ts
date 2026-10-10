import * as v from "valibot"

/** One document library (= Graph drive), addressed in paths as `<siteKey>/<name>/…`. */
export const sharepointLibrarySchema = v.object({
  siteKey: v.string(),
  siteName: v.string(),
  siteUrl: v.string(),
  name: v.string(),
  driveId: v.string(),
  driveUrl: v.string(),
})

export type SharepointLibrary = v.InferOutput<typeof sharepointLibrarySchema>
