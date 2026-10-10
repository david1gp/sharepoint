import { expect, test } from "bun:test"
import { sharepointItemPathEncode } from "../src/sharepointItemPathEncode.js"
import type { SharepointLibrary } from "../src/sharepointLibrarySchema.js"
import { sharepointPathFormat } from "../src/sharepointPathFormat.js"
import { sharepointPathParse } from "../src/sharepointPathParse.js"
import { sharepointSearchRegionRead } from "../src/sharepointSearchRegionRead.js"
import { sharepointSiteKeysCreate } from "../src/sharepointSiteKeysCreate.js"

const library = (siteKey: string, name: string, driveId: string): SharepointLibrary => ({
  siteKey,
  siteName: siteKey,
  siteUrl: `https://contoso.sharepoint.com/sites/${siteKey}`,
  name,
  driveId,
  driveUrl: `https://contoso.sharepoint.com/sites/${siteKey}/${name}`,
})
const libraries = [library("Kunden", "Dokumente", "d1"), library("park-management", "Template", "d2")]

test("site keys use the last URL segment, 'root' for the tenant root, and full paths only when names collide", () => {
  const keys = sharepointSiteKeysCreate([
    "https://contoso.sharepoint.com/",
    "https://contoso.sharepoint.com/sites/M%C3%BCnchen",
    "https://contoso.sharepoint.com/sites/team",
    "https://contoso.sharepoint.com/teams/team",
  ])
  expect([...keys.values()]).toEqual(["root", "München", "sites-team", "teams-team"])
})

test("paths resolve site and library case-insensitively and keep the path inside the library", () => {
  const parsed = sharepointPathParse("kunden/DOKUMENTE/Test Ordner/Rüge.pdf", libraries)
  expect(parsed.success && parsed.data.library.driveId).toBe("d1")
  expect(parsed.success && parsed.data.itemPath).toBe("Test Ordner/Rüge.pdf")
  const folder = sharepointPathParse("Kunden/Dokumente/Test Ordner/", libraries)
  expect(folder.success && folder.data.trailingSlash).toBe(true)
})

test("with a configured site, paths start at the library and output omits the site", () => {
  const parsed = sharepointPathParse("Dokumente/a.pdf", libraries, { site: "Kunden" })
  expect(parsed.success && parsed.data.itemPath).toBe("a.pdf")
  expect(sharepointPathFormat(libraries[0]!, "a.pdf", { site: "Kunden" })).toBe("Dokumente/a.pdf")
  expect(sharepointPathFormat(libraries[0]!, "Ordner", { folder: true })).toBe("Kunden/Dokumente/Ordner/")
})

test("unknown sites or libraries and '..' segments are rejected, listing the valid choices", () => {
  const site = sharepointPathParse("Nope/Dokumente", libraries)
  expect(!site.success && site.errorMessage).toContain("Kunden, park-management")
  const lib = sharepointPathParse("Kunden/Nope", libraries)
  expect(!lib.success && lib.errorMessage).toContain("Dokumente")
  expect(sharepointPathParse("Kunden/Dokumente/../x", libraries).success).toBe(false)
  expect(sharepointPathParse("Kunden", libraries).success).toBe(false)
})

test("item paths are encoded per segment so spaces, umlauts, # and ? reach Graph intact", () => {
  expect(sharepointItemPathEncode("Test Ordner/Rüge #1?.pdf")).toBe("Test%20Ordner/R%C3%BCge%20%231%3F.pdf")
})

test("the search region is read from Graph's 'valid regions' error message", () => {
  expect(sharepointSearchRegionRead("Requested region  not found. Only valid regions are EMEA.")).toBe("EMEA")
  expect(sharepointSearchRegionRead("Only valid regions are NAM, EUR")).toBe("NAM")
  expect(sharepointSearchRegionRead("Region is required when request with application permission.")).toBeUndefined()
})
