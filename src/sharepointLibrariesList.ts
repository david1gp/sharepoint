import type { Result } from "#result"
import { createResult, createResultError } from "#result"
import { sharepointErrorRead } from "./sharepointErrorRead.js"
import type { SharepointGraph } from "./sharepointGraphCreate.js"
import type { SharepointLibrary } from "./sharepointLibrarySchema.js"
import { sharepointSiteKeysCreate } from "./sharepointSiteKeysCreate.js"

type SharepointLibrariesListOptions = Readonly<{
  /** Discovered from the tenant root site when omitted. */
  hostname?: string
  /** Only this site (matched case-insensitively against its key). */
  site?: string
  /** Called for sites whose libraries cannot be read; listing continues. */
  warn?: (message: string) => void
}>

type Site = { id: string; webUrl: string; displayName: string }

/** All document libraries of all sites on the tenant host (personal OneDrives and hidden containers excluded). */
export async function sharepointLibrariesList(
  graph: SharepointGraph,
  options: SharepointLibrariesListOptions = {},
): Promise<Result<SharepointLibrary[]>> {
  const op = "sharepointLibrariesList"
  try {
    let hostname = options.hostname
    if (!hostname) {
      const root = await graph.sites.root.select("webUrl")()
      if (!root.webUrl) return createResultError(op, "Tenant root site has no webUrl; set hostname in the config")
      hostname = new URL(root.webUrl).hostname
    }
    const query = graph.sites.select("id", "webUrl", "displayName")
    query.query.set("search", "*")
    const sites: Site[] = []
    for await (const page of query) {
      for (const site of page) {
        if (!site.id || !site.webUrl) continue
        const url = new URL(site.webUrl)
        // contentstorage/ holds Loop/Embedded containers: not user sites, and app-only access is denied anyway
        if (url.hostname !== hostname || url.pathname.startsWith("/contentstorage/")) continue
        sites.push({ id: site.id, webUrl: site.webUrl, displayName: site.displayName ?? site.webUrl })
      }
    }
    const keys = sharepointSiteKeysCreate(sites.map((site) => site.webUrl))
    const selected = options.site
      ? sites.filter((site) => keys.get(site.webUrl)?.toLowerCase() === options.site?.toLowerCase())
      : sites
    if (options.site && selected.length === 0) {
      const available = [...keys.values()].sort().join(", ")
      return createResultError(op, `Configured site "${options.site}" not found on ${hostname}. Sites: ${available}`)
    }
    const libraries: SharepointLibrary[] = []
    for (const site of selected) {
      const drives = await graph.sites
        .getById(site.id)
        .drives.select("id", "name", "webUrl")()
        .catch((error: unknown) => {
          options.warn?.(`skipping ${site.webUrl}: ${sharepointErrorRead(error).message}`)
          return []
        })
      for (const drive of drives) {
        if (!drive.id || !drive.name || !drive.webUrl) continue
        libraries.push({
          siteKey: keys.get(site.webUrl) ?? site.webUrl,
          siteName: site.displayName,
          siteUrl: site.webUrl,
          name: drive.name,
          driveId: drive.id,
          driveUrl: drive.webUrl,
        })
      }
    }
    return createResult(libraries.sort((a, b) => a.siteKey.localeCompare(b.siteKey) || a.name.localeCompare(b.name)))
  } catch (error) {
    return createResultError(op, sharepointErrorRead(error).message, String(error))
  }
}
