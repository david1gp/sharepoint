import type { SharepointLibrary } from "./sharepointLibrarySchema.js"

/** Inverse of sharepointPathParse: the user-facing path, without the site when one is configured. */
export function sharepointPathFormat(
  library: SharepointLibrary,
  itemPath: string,
  options: Readonly<{ site?: string; folder?: boolean }> = {},
): string {
  const parts = options.site ? [library.name, itemPath] : [library.siteKey, library.name, itemPath]
  const path = parts.filter(Boolean).join("/")
  return options.folder ? `${path}/` : path
}
