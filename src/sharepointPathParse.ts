import type { Result } from "#result"
import { createResult, createResultError } from "#result"
import type { SharepointLibrary } from "./sharepointLibrarySchema.js"

export type SharepointPath = Readonly<{ library: SharepointLibrary; itemPath: string; trailingSlash: boolean }>

/**
 * `<site>/<library>/<path>` → library + path inside it (site/library match case-insensitively).
 * With a configured `site`, paths are `<library>/<path>` instead.
 */
export function sharepointPathParse(
  path: string,
  libraries: readonly SharepointLibrary[],
  options: Readonly<{ site?: string }> = {},
): Result<SharepointPath> {
  const op = "sharepointPathParse"
  const segments = path.split("/").filter(Boolean)
  if (segments.some((segment) => segment === "." || segment === ".."))
    return createResultError(op, `"." and ".." are not allowed in "${path}"`)
  const [siteKey, libraryName, ...rest] = options.site ? [options.site, ...segments] : segments
  const expected = options.site ? "<library>[/path]" : "<site>/<library>[/path]"
  if (!siteKey || !libraryName) return createResultError(op, `Expected ${expected}, got "${path}"`)
  const siteLibraries = libraries.filter((library) => library.siteKey.toLowerCase() === siteKey.toLowerCase())
  if (siteLibraries.length === 0) {
    const sites = [...new Set(libraries.map((library) => library.siteKey))].join(", ")
    return createResultError(op, `Unknown site "${siteKey}". Sites: ${sites}`)
  }
  const library = siteLibraries.find((candidate) => candidate.name.toLowerCase() === libraryName.toLowerCase())
  if (!library) {
    const names = siteLibraries.map((candidate) => candidate.name).join(", ")
    return createResultError(op, `Unknown library "${libraryName}" on ${siteKey}. Libraries: ${names}`)
  }
  return createResult({ library, itemPath: rest.join("/"), trailingSlash: path.endsWith("/") })
}
