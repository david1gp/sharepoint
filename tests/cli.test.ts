import { afterAll, expect, test } from "bun:test"
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import { join } from "node:path"
import pkg from "../package.json" with { type: "json" }
import { graphFake } from "./graphFake.js"

const fake = graphFake()
const directory = await mkdtemp(join(tmpdir(), "sharepoint-cli-test-"))
const configPath = join(directory, "config.json")
await writeFile(
  configPath,
  JSON.stringify({
    tenantId: "00000000-0000-0000-0000-000000000001",
    clientId: "00000000-0000-0000-0000-000000000002",
    clientSecret: "abc~defghijklmnopqrstuvwxyz",
  }),
)
afterAll(async () => {
  fake.stop()
  await rm(directory, { recursive: true, force: true })
})

async function cli(args: readonly string[], env: Record<string, string> = {}) {
  const child = Bun.spawn(["bun", "run", `${import.meta.dir}/../src/cli.ts`, ...args], {
    cwd: directory,
    env: { HOME: directory, PATH: process.env.PATH ?? "", SHAREPOINT_GRAPH_URL: `${fake.url}/`, ...env },
    stdout: "pipe",
    stderr: "pipe",
  })
  const [stdout, stderr] = await Promise.all([new Response(child.stdout).text(), new Response(child.stderr).text()])
  return { exitCode: await child.exited, stdout, stderr }
}

test("--version prints the package version", async () => {
  expect((await cli(["--version"])).stdout).toBe(`${pkg.version}\n`)
})

test("--config, ls and get - work end to end against Graph", async () => {
  const ls = await cli(["ls", "--config", configPath, "Kunden/Dokumente/Test Ordner"])
  expect(ls.stdout).toBe("Kunden/Dokumente/Test Ordner/Rüge.pdf\n")
  const get = await cli(["get", "Kunden/Dokumente/Test Ordner/Rüge.pdf", "-"], { SHAREPOINT_CONFIG: configPath })
  expect(get.stdout).toBe("%PDF-rüge")
})

test("put uploads a local file into a folder and refuses to overwrite it again", async () => {
  await writeFile(join(directory, "bericht.txt"), "inhalt")
  const env = { SHAREPOINT_CONFIG: configPath }
  expect((await cli(["put", "bericht.txt", "Verwaltung/Documents/cli/"], env)).exitCode).toBe(0)
  const again = await cli(["put", "bericht.txt", "Verwaltung/Documents/cli/"], env)
  expect(again.exitCode).toBe(1)
  expect(again.stderr).toContain("already exists; use --overwrite")
  await cli(["get", "Verwaltung/Documents/cli/bericht.txt", "copy.txt"], env)
  expect(await readFile(join(directory, "copy.txt"), "utf8")).toBe("inhalt")
})

test("search --json returns hits with downloadable paths", async () => {
  const search = await cli(["search", "--json", "rüge"], { SHAREPOINT_CONFIG: configPath })
  expect(JSON.parse(search.stdout).hits[0].path).toBe("Kunden/Dokumente/Test Ordner/Rüge.pdf")
})

test("invalid arguments exit with usage code 2, missing config with 1 and a hint", async () => {
  expect((await cli(["search", "--limit", "0", "x"], { SHAREPOINT_CONFIG: configPath })).exitCode).toBe(2)
  expect((await cli(["ls", "Kunden/../x"], { SHAREPOINT_CONFIG: configPath })).exitCode).toBe(2)
  const missing = await cli(["ls"])
  expect(missing.exitCode).toBe(1)
  expect(missing.stderr).toContain(".config/sharepoint/config.json")
})

test("SHAREPOINT_GRAPH_URL only accepts a local test server", async () => {
  const remote = await cli(["ls"], { SHAREPOINT_CONFIG: configPath, SHAREPOINT_GRAPH_URL: "https://evil.example/" })
  expect(remote.exitCode).toBe(1)
  expect(remote.stderr).toContain("must be a local")
})
