import { sharepointUrlDecode } from "./sharepointUrlDecode.js"

/**
 * Short, path-safe site names: the last URL segment (`/sites/park-management` → `park-management`, `/` → `root`).
 * When two sites share a last segment, both fall back to their full path joined with "-" (`sites-a`, `teams-a`).
 */
export function sharepointSiteKeysCreate(webUrls: readonly string[]): Map<string, string> {
  const segmentsOf = (webUrl: string) => sharepointUrlDecode(new URL(webUrl).pathname).split("/").filter(Boolean)
  const shortKey = (webUrl: string) => segmentsOf(webUrl).at(-1) ?? "root"
  const counts = new Map<string, number>()
  for (const webUrl of webUrls) counts.set(shortKey(webUrl), (counts.get(shortKey(webUrl)) ?? 0) + 1)
  const keys = new Map<string, string>()
  for (const webUrl of webUrls) {
    const key = shortKey(webUrl)
    keys.set(webUrl, (counts.get(key) ?? 0) > 1 ? segmentsOf(webUrl).join("-") || "root" : key)
  }
  return keys
}
