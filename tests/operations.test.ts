import { afterAll, expect, test } from "bun:test"
import { sharepointClientCreate } from "../src/sharepointClientCreate.js"
import type { SharepointConfig } from "../src/sharepointConfigSchema.js"
import { sharepointDelete } from "../src/sharepointDelete.js"
import { sharepointDownload } from "../src/sharepointDownload.js"
import { sharepointFolderCreate } from "../src/sharepointFolderCreate.js"
import { sharepointList } from "../src/sharepointList.js"
import { sharepointSearch } from "../src/sharepointSearch.js"
import { sharepointUpload } from "../src/sharepointUpload.js"
import { graphFake } from "./graphFake.js"

const fake = graphFake()
afterAll(() => fake.stop())
const config: SharepointConfig = {
  tenantId: "00000000-0000-0000-0000-000000000001",
  clientId: "00000000-0000-0000-0000-000000000002",
  clientSecret: "abc~defghijklmnopqrstuvwxyz",
}
const client = async (overrides: Partial<SharepointConfig> = {}) => {
  const created = await sharepointClientCreate({ ...config, ...overrides }, { graph: fake.graph })
  if (!created.success) throw new Error(created.errorMessage)
  return created.data
}
const text = (value: string) => new TextEncoder().encode(value)

test("libraries are discovered for all sites on the root host, skipping hidden contentstorage containers", async () => {
  const sp = await client()
  expect(sp.libraries.map((library) => `${library.siteKey}/${library.name}`)).toEqual([
    "Kunden/Dokumente",
    "Verwaltung/Documents",
  ])
  const sites = await sharepointList(sp)
  expect(sites.success && sites.data.map((entry) => entry.path)).toEqual(["Kunden/", "Verwaltung/"])
})

test("listing a folder returns <site>/<library>/<path> entries that get/rm accept", async () => {
  const listed = await sharepointList(await client(), "kunden/dokumente/Test Ordner")
  expect(listed.success && listed.data).toMatchObject([{ path: "Kunden/Dokumente/Test Ordner/Rüge.pdf", type: "file" }])
})

test("download returns file contents and refuses folders", async () => {
  const sp = await client()
  const file = await sharepointDownload(sp, "Kunden/Dokumente/Test Ordner/Rüge.pdf")
  expect(file.success && new TextDecoder().decode(file.data.content)).toBe("%PDF-rüge")
  expect((await sharepointDownload(sp, "Kunden/Dokumente/Test Ordner")).success).toBe(false)
})

test("upload never replaces an existing file unless overwrite is set", async () => {
  const sp = await client()
  const path = "Verwaltung/Documents/neu/"
  expect((await sharepointUpload(sp, { path, fileName: "a.txt", content: text("v1") })).success).toBe(true)
  const again = await sharepointUpload(sp, { path, fileName: "a.txt", content: text("v2") })
  expect(!again.success && again.errorMessage).toContain("already exists")
  expect(fake.drives[1]?.files.get("neu/a.txt")?.content).toBe("v1")
  await sharepointUpload(sp, { path, fileName: "a.txt", content: text("v2"), overwrite: true })
  expect(fake.drives[1]?.files.get("neu/a.txt")?.content).toBe("v2")
})

test("mkdir creates missing parents and keeps existing folders", async () => {
  const sp = await client()
  expect((await sharepointFolderCreate(sp, "Verwaltung/Documents/x/y/z")).success).toBe(true)
  expect((await sharepointFolderCreate(sp, "Verwaltung/Documents/x/y")).success).toBe(true)
  expect([...(fake.drives[1]?.folders ?? [])]).toEqual(expect.arrayContaining(["x", "x/y", "x/y/z"]))
})

test("delete refuses folders without recursive and never deletes a whole library", async () => {
  const sp = await client()
  await sharepointFolderCreate(sp, "Verwaltung/Documents/tmp")
  expect((await sharepointDelete(sp, { path: "Verwaltung/Documents/tmp" })).success).toBe(false)
  expect((await sharepointDelete(sp, { path: "Verwaltung/Documents/tmp", recursive: true })).success).toBe(true)
  expect((await sharepointDelete(sp, { path: "Verwaltung/Documents", recursive: true })).success).toBe(false)
})

test("search discovers the region from Graph and maps hits back to CLI paths", async () => {
  const result = await sharepointSearch(await client(), { query: "rüge" })
  expect(result.success && result.data.region).toBe("EMEA")
  expect(result.success && result.data.hits.map((hit) => hit.path)).toEqual(["Kunden/Dokumente/Test Ordner/Rüge.pdf"])
  expect(result.success && result.data.hits[0]?.summary).toBe("rüge match …")
})

test("a configured site scopes discovery and paths to <library>/<path>", async () => {
  const sp = await client({ site: "kunden" })
  expect(sp.libraries.map((library) => library.name)).toEqual(["Dokumente"])
  const listed = await sharepointList(sp, "Dokumente/Test Ordner")
  expect(listed.success && listed.data[0]?.path).toBe("Dokumente/Test Ordner/Rüge.pdf")
  const missing = await sharepointClientCreate({ ...config, site: "nope" }, { graph: fake.graph })
  expect(!missing.success && missing.errorMessage).toContain("Kunden, Verwaltung")
})
