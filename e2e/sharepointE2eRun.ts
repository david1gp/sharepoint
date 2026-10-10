#!/usr/bin/env bun
/**
 * Live end-to-end test against a real tenant (`bun run test:e2e`; never part of `bun test`).
 * Uses the normal config (~/.config/sharepoint/config.json, SHAREPOINT_CONFIG or SHAREPOINT_* variables).
 *
 *   SHAREPOINT_E2E_PATH    writable folder for the test folder   (default: Verwaltung/Documents)
 *   SHAREPOINT_E2E_QUERY   term that already matches indexed files (default: Rüge)
 *   --wait-index           also wait (≤10 min) until search finds the freshly uploaded file
 *
 * Everything is created inside `sharepoint-davids-e2e-test-<timestamp>` and that folder is deleted (to the
 * site recycle bin) at the end, also when a step fails.
 */
import { sharepointClientCreate } from "../src/sharepointClientCreate.js"
import { sharepointConfigLoad } from "../src/sharepointConfigLoad.js"
import { sharepointDelete } from "../src/sharepointDelete.js"
import { sharepointDownload } from "../src/sharepointDownload.js"
import { sharepointFolderCreate } from "../src/sharepointFolderCreate.js"
import { sharepointList } from "../src/sharepointList.js"
import { sharepointSearch } from "../src/sharepointSearch.js"
import { sharepointUpload } from "../src/sharepointUpload.js"

const basePath = (process.env.SHAREPOINT_E2E_PATH ?? "Verwaltung/Documents").replace(/\/+$/, "")
const knownQuery = process.env.SHAREPOINT_E2E_QUERY ?? "Rüge"
const waitIndex = process.argv.includes("--wait-index")
const stamp = new Date().toISOString().replace(/[-:]/g, "").replace(/\..*/, "").replace("T", "-")
const testFolder = `${basePath}/sharepoint-davids-e2e-test-${stamp}`
const marker = `sharepointe2emarker${stamp.replace("-", "")}`
const encoder = new TextEncoder()
const decoder = new TextDecoder()

let failures = 0
async function step(name: string, run: () => Promise<string | undefined>) {
  const started = performance.now()
  try {
    const detail = await run()
    console.log(`✓ ${name} (${Math.round(performance.now() - started)} ms)${detail ? ` - ${detail}` : ""}`)
  } catch (error) {
    failures++
    console.error(`✗ ${name}: ${error instanceof Error ? error.message : String(error)}`)
  }
}
function unwrap<T>(result: { success: true; data: T } | { success: false; errorMessage: string }): T {
  if (!result.success) throw new Error(result.errorMessage)
  return result.data
}
function expectEqual(actual: unknown, expected: unknown, label: string) {
  if (actual !== expected)
    throw new Error(`${label}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`)
}

const config = await sharepointConfigLoad()
if (!config.success) {
  console.error(`config: ${config.errorMessage}`)
  process.exit(1)
}
const created = await sharepointClientCreate(config.data, { warn: (message) => console.warn(`  warning: ${message}`) })
if (!created.success) {
  console.error(`sign-in / discovery: ${created.errorMessage}`)
  process.exit(1)
}
const client = created.data
console.log(`tenant libraries: ${client.libraries.length}, test folder: ${testFolder}\n`)

const fileName = "e2e ä ö ü #1.txt"
const filePath = `${testFolder}/sub folder/${fileName}`
const content = `sharepoint e2e ${marker}\nümlauts ✓\n`

try {
  await step("list sites/libraries", async () => {
    const entries = unwrap(await sharepointList(client))
    if (entries.length === 0) throw new Error("no sites or libraries visible")
    return `${entries.length} entries`
  })
  await step("list base folder", async () => `${unwrap(await sharepointList(client, basePath)).length} entries`)
  await step(`search existing content "${knownQuery}"`, async () => {
    const result = unwrap(await sharepointSearch(client, { query: knownQuery, limit: 5 }))
    const hit = result.hits.find((candidate) => !candidate.path.startsWith("http"))
    if (!hit) throw new Error(`no downloadable hit (total ${result.total}); set SHAREPOINT_E2E_QUERY`)
    const file = unwrap(await sharepointDownload(client, hit.path))
    if (file.content.byteLength === 0) throw new Error(`${hit.path} downloaded empty`)
    return `${result.total} hit(s), region ${result.region}, read ${hit.path} (${file.content.byteLength} B)`
  })
  await step("mkdir (with parents)", async () =>
    unwrap(await sharepointFolderCreate(client, `${testFolder}/sub folder`)),
  )
  await step("upload", async () => {
    const uploaded = unwrap(
      await sharepointUpload(client, { path: `${testFolder}/sub folder/`, fileName, content: encoder.encode(content) }),
    )
    return uploaded.path
  })
  await step("list uploaded file", async () => {
    const entries = unwrap(await sharepointList(client, `${testFolder}/sub folder`))
    expectEqual(
      entries.find((entry) => entry.path.endsWith(fileName))?.size,
      encoder.encode(content).byteLength,
      "size",
    )
    return undefined
  })
  await step("download and compare", async () => {
    expectEqual(decoder.decode(unwrap(await sharepointDownload(client, filePath)).content), content, "content")
    return undefined
  })
  await step("upload without --overwrite is refused", async () => {
    const again = await sharepointUpload(client, { path: filePath, fileName, content: encoder.encode("must not land") })
    if (again.success) throw new Error("existing file was replaced without overwrite")
    expectEqual(decoder.decode(unwrap(await sharepointDownload(client, filePath)).content), content, "content kept")
    return undefined
  })
  await step("upload with overwrite replaces", async () => {
    unwrap(await sharepointUpload(client, { path: filePath, fileName, content: encoder.encode("v2"), overwrite: true }))
    expectEqual(decoder.decode(unwrap(await sharepointDownload(client, filePath)).content), "v2", "content")
    unwrap(
      await sharepointUpload(client, { path: filePath, fileName, content: encoder.encode(content), overwrite: true }),
    )
    return undefined
  })
  await step("delete folder needs recursive", async () => {
    if ((await sharepointDelete(client, { path: testFolder })).success)
      throw new Error("folder deleted without recursive")
    return undefined
  })
  if (waitIndex)
    await step("search finds the uploaded file (indexing)", async () => {
      const deadline = Date.now() + 10 * 60 * 1000
      while (Date.now() < deadline) {
        const result = unwrap(await sharepointSearch(client, { query: marker, limit: 5 }))
        const hit = result.hits.find((candidate) => candidate.path.endsWith(fileName))
        if (hit) return hit.path
        await Bun.sleep(20_000)
      }
      throw new Error("not indexed within 10 minutes")
    })
} finally {
  await step("cleanup: delete test folder", async () =>
    unwrap(await sharepointDelete(client, { path: testFolder, recursive: true })),
  )
  await step("cleanup: test folder is gone", async () => {
    if ((await sharepointList(client, testFolder)).success) throw new Error(`${testFolder} still exists`)
    return undefined
  })
}

console.log(failures === 0 ? "\nall e2e steps passed" : `\n${failures} e2e step(s) failed`)
process.exit(failures === 0 ? 0 : 1)
