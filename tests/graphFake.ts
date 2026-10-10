import { graphfi } from "@pnp/graph"
import "@pnp/graph/sites/index.js"
import "@pnp/graph/files/index.js"
import "@pnp/graph/search/index.js"
import { GraphDefault } from "@pnp/nodejs"
import type { SharepointGraph } from "../src/sharepointGraphCreate.js"

type FakeFile = { content: string }
type FakeDrive = { id: string; name: string; webUrl: string; files: Map<string, FakeFile>; folders: Set<string> }

/**
 * Minimal in-memory Microsoft Graph: two sites, path addressing, upload/download, folders, delete and search
 * (region required, like app-only Graph). Returns an unauthenticated PnP client pointed at it.
 */
export function graphFake() {
  const host = "https://contoso.sharepoint.com"
  const drives: FakeDrive[] = [
    {
      id: "d-kunden",
      name: "Dokumente",
      webUrl: `${host}/Kunden/Freigegebene%20Dokumente`,
      files: new Map([["Test Ordner/Rüge.pdf", { content: "%PDF-rüge" }]]),
      folders: new Set(["Test Ordner"]),
    },
    {
      id: "d-verwaltung",
      name: "Documents",
      webUrl: `${host}/sites/Verwaltung/Shared%20Documents`,
      files: new Map(),
      folders: new Set(),
    },
  ]
  const sites = [
    { id: "s-root", webUrl: `${host}/`, displayName: "Root", drives: [] as FakeDrive[] },
    { id: "s-kunden", webUrl: `${host}/Kunden`, displayName: "Kunden", drives: [drives[0]!] },
    { id: "s-verwaltung", webUrl: `${host}/sites/Verwaltung`, displayName: "Verwaltung", drives: [drives[1]!] },
    { id: "s-hidden", webUrl: `${host}/contentstorage/x`, displayName: "hidden", drives: [] as FakeDrive[] },
  ]
  const json = (body: unknown, status = 200) => Response.json(body, { status })
  const notFound = () => json({ error: { code: "itemNotFound", message: "The resource could not be found." } }, 404)
  const item = (drive: FakeDrive, path: string) => {
    const file = drive.files.get(path)
    const name = path.split("/").at(-1) ?? ""
    const webUrl = `${drive.webUrl}/${path.split("/").map(encodeURIComponent).join("/")}`
    if (file)
      return {
        id: `${drive.id}:${path}`,
        name,
        size: file.content.length,
        file: {},
        webUrl,
        "@microsoft.graph.downloadUrl": `${origin()}/download/${drive.id}/${encodeURIComponent(path)}`,
      }
    if (path === "" || drive.folders.has(path)) return { id: `${drive.id}:${path}`, name, folder: {}, webUrl }
    return undefined
  }
  const children = (drive: FakeDrive, folder: string) => {
    const prefix = folder ? `${folder}/` : ""
    const direct = (path: string) => path.startsWith(prefix) && !path.slice(prefix.length).includes("/")
    return [...drive.folders, ...drive.files.keys()].filter(direct).map((path) => item(drive, path))
  }
  let server: ReturnType<typeof Bun.serve> | undefined
  const origin = () => `http://localhost:${server?.port}`

  async function handle(request: Request): Promise<Response> {
    const url = new URL(request.url)
    const path = decodeURIComponent(url.pathname.replace(/^\/v1\.0/, "").replace(/\/\//g, "/"))
    const download = /^\/download\/([^/]+)\/(.+)$/.exec(path)
    if (download) {
      const file = drives.find((drive) => drive.id === download[1])?.files.get(download[2] ?? "")
      return file ? new Response(file.content) : notFound()
    }
    if (path === "/sites/root") return json({ webUrl: `${host}/` })
    if (path === "/sites") return json({ value: sites.map(({ drives: _, ...site }) => site) })
    const siteDrives = /^\/sites\/([^/]+)\/drives$/.exec(path)
    if (siteDrives) return json({ value: sites.find((site) => site.id === siteDrives[1])?.drives ?? [] })
    if (path === "/search/query" && request.method === "POST") {
      const body = (await request.json()) as { requests: { region?: string; query: { queryString: string } }[] }
      const search = body.requests[0]
      if (search?.region !== "EMEA")
        return json({ error: { message: "Requested region  not found. Only valid regions are EMEA." } }, 400)
      const term = search.query.queryString.split(" ")[0]?.toLowerCase() ?? ""
      const hits = drives.flatMap((drive) =>
        [...drive.files.keys()]
          .filter((filePath) => filePath.toLowerCase().includes(term))
          .map((filePath) => ({
            summary: `<c0>${term}</c0> match<ddd/>`,
            resource: { ...item(drive, filePath), parentReference: { driveId: drive.id } },
          })),
      )
      return json({ value: [{ hitsContainers: [{ total: hits.length, hits }] }] })
    }
    const driveMatch = /^\/drives\/([^/]+)\/(root.*)$/.exec(path)
    const drive = drives.find((candidate) => candidate.id === driveMatch?.[1])
    if (!driveMatch || !drive) return notFound()
    const rest = driveMatch[2] ?? ""
    const byPath = /^root:\/(.+?):(\/children|\/content)?$/.exec(rest)
    const itemPath = byPath?.[1] ?? ""
    const suffix = byPath ? (byPath[2] ?? "") : rest.replace(/^root/, "")
    if (suffix === "/content" && request.method === "PUT") {
      const parent = itemPath.split("/").slice(0, -1)
      for (let index = 1; index <= parent.length; index++) drive.folders.add(parent.slice(0, index).join("/"))
      drive.files.set(itemPath, { content: await request.text() })
      return json(item(drive, itemPath), 201)
    }
    if (suffix === "/children" && request.method === "POST") {
      const body = (await request.json()) as { name: string }
      const folderPath = [itemPath, body.name].filter(Boolean).join("/")
      if (item(drive, folderPath)) return json({ error: { message: "Name already exists" } }, 409)
      drive.folders.add(folderPath)
      return json(item(drive, folderPath), 201)
    }
    if (suffix === "/children") return item(drive, itemPath) ? json({ value: children(drive, itemPath) }) : notFound()
    if (request.method === "DELETE") {
      if (!item(drive, itemPath)) return notFound()
      for (const key of [...drive.files.keys()])
        if (key === itemPath || key.startsWith(`${itemPath}/`)) drive.files.delete(key)
      for (const key of [...drive.folders])
        if (key === itemPath || key.startsWith(`${itemPath}/`)) drive.folders.delete(key)
      return new Response(null, { status: 204 })
    }
    const found = item(drive, itemPath)
    return found ? json(found) : notFound()
  }

  server = Bun.serve({ port: 0, fetch: handle })
  const baseUrl = `http://localhost:${server.port}/v1.0`
  const graph = graphfi().using(GraphDefault({ baseUrl })) as SharepointGraph
  return {
    graph,
    drives,
    url: baseUrl,
    stop: () => server?.stop(true),
  }
}
