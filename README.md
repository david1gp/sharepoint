# @adaptive-ds/sharepoint

SharePoint files from the terminal and from TypeScript. List, search, download, upload, create folders and delete across **all sites and document libraries** of a Microsoft 365 tenant, using one short path syntax:

```sh
sharepoint ls                                         # sites
sharepoint search "invoice filetype:pdf"              # full-text search across every site
sharepoint get "Sales/Documents/Offers/2026-001.pdf"  # → ./2026-001.pdf
sharepoint put report.pdf "Sales/Documents/Reports/"  # never replaces existing files by accident
```

It uses Microsoft Graph with **app-only authentication** (an Entra app registration and a client secret), so it works headless on servers, in CI and for automation or coding agents. There is no interactive login and no certificate. Every library function returns a `Result` instead of throwing, and config and arguments are validated with [valibot](https://valibot.dev).

## Requirements

- [Bun](https://bun.sh/) `>= 1.3` or Node.js `>= 22`
- An Entra app registration with the Microsoft Graph **application** permission `Sites.ReadWrite.All` (or `Sites.Read.All`, or `Sites.Selected` for chosen sites) and admin consent (see [Setup](#setup))

## Installation

```sh
bun add --global @adaptive-ds/sharepoint @adaptive-ds/result valibot   # the `sharepoint` command
bun add @adaptive-ds/sharepoint @adaptive-ds/result valibot             # the library
```

## Configuration

Create `~/.config/sharepoint/config.json`, readable only by you (`chmod 600`):

```json
{
  "tenantId": "00000000-0000-0000-0000-000000000000",
  "clientId": "00000000-0000-0000-0000-000000000000",
  "clientSecret": "the secret *value*"
}
```

| Field | Required | Meaning |
| --- | --- | --- |
| `tenantId` | yes | Directory (tenant) ID |
| `clientId` | yes | Application (client) ID |
| `clientSecret` | yes | Client secret **value** (not the secret ID) |
| `site` | no | Work with one site only; paths become `<library>/<path>` |
| `hostname` | no | e.g. `contoso.sharepoint.com`; discovered from the tenant root site by default |
| `searchRegion` | no | Microsoft Search region (e.g. `EMEA`, `NAM`); discovered automatically by default |

Configuration sources, highest priority first:

1. `--config <file>`
2. `SHAREPOINT_CONFIG=<file>`
3. `SHAREPOINT_TENANT_ID`, `SHAREPOINT_CLIENT_ID`, `SHAREPOINT_CLIENT_SECRET`, `SHAREPOINT_SITE`, `SHAREPOINT_HOSTNAME` and `SHAREPOINT_SEARCH_REGION` (these override single fields of the file)
4. `~/.config/sharepoint/config.json`

## Paths

`<site>/<library>/<folder>/<file>`. Site and library names match case-insensitively. Use `sharepoint ls` and `sharepoint ls <site>` to see the valid names.

- A site name is the last part of its URL: `…/sites/park-management` → `park-management`. The tenant root site is `root`.
- A trailing `/` on an upload target means "into this folder", and the local file name is kept.
- With `site` configured, leave the site out: `Documents/Reports/q3.pdf`.

## CLI

```text
sharepoint ls [--long] [--json] [PATH]
sharepoint get PATH [TARGET | -]
sharepoint put [--overwrite] FILE PATH
sharepoint mkdir PATH
sharepoint rm [--recursive] PATH
sharepoint search [--limit COUNT] [--json] QUERY
```

| Command | Does |
| --- | --- |
| `ls` | Lists sites (no path), the libraries of a site, or a folder. `-l` adds size and modified time; `--json` is machine-readable. |
| `get` | Downloads to the file name, to `TARGET`, or to stdout with `-`. |
| `put` | Uploads a file. Missing folders are created. An existing file is **only** replaced with `--overwrite`. |
| `mkdir` | Creates a folder and missing parents (like `mkdir -p`). |
| `rm` | Moves a file to the site's **recycle bin**. Folders need `-r`. Refuses to delete a whole library. |
| `search` | Microsoft Search across all files, including content. Supports [KQL](https://learn.microsoft.com/sharepoint/dev/general-development/keyword-query-language-kql-syntax-reference): words, `"phrases"`, `filetype:xlsx`, `path:"https://…"`. Results are paths you can pass to `get`. |

Exit codes: `0` success, `1` failure (auth, not found, refused), `2` invalid arguments. Every command also accepts `--config`, `--help` and `--version`.

For bulk copies and syncing, [rclone's onedrive backend](https://rclone.org/onedrive/) works with the same app registration (`client_credentials = true`, `drive_type = documentlibrary`).

## Library

```ts
import { sharepointClientCreate, sharepointConfigLoad, sharepointList, sharepointUpload } from "@adaptive-ds/sharepoint"

const config = await sharepointConfigLoad()
if (!config.success) throw new Error(config.errorMessage)
const client = await sharepointClientCreate(config.data)
if (!client.success) throw new Error(client.errorMessage)

const files = await sharepointList(client.data, "Sales/Documents/Reports")
const uploaded = await sharepointUpload(client.data, {
  path: "Sales/Documents/Reports/",
  fileName: "q3.pdf",
  content: await Bun.file("q3.pdf").bytes(),
})
```

Also exported: `sharepointDownload`, `sharepointFolderCreate`, `sharepointDelete`, `sharepointSearch`, `sharepointPathParse` and `sharepointGraphCreate` (a configured [PnPjs](https://pnp.github.io/pnpjs/) Graph client for anything else). The valibot schemas and their types are exported too.

## Setup

In the [Entra admin center](https://entra.microsoft.com):

1. Go to **App registrations → New registration**. Choose single tenant and no redirect URI. The *Overview* page shows `clientId` and `tenantId`.
2. Go to **Certificates & secrets → Client secrets → New client secret**. Copy the **Value** into `clientSecret`.
3. Go to **API permissions → Add a permission → Microsoft Graph → Application permissions** and add `Sites.ReadWrite.All` (or `Sites.Read.All`). Then click **Grant admin consent**.
4. Run `sharepoint ls`.

To limit access to selected sites, use `Sites.Selected` instead. Then grant each site with `POST /v1.0/sites/{site-id}/permissions`, giving `roles: ["write"]` and `grantedToIdentities` set to your app.

Why Graph and not SharePoint's own REST API (`@pnp/sp`)? SharePoint REST refuses client secrets for app-only tokens and requires a certificate. Graph covers files, folders, libraries and search with a plain secret.

## Development

```sh
bun install
bun run test        # unit + CLI tests against an in-memory Graph fake
bun run test:e2e    # live test against your tenant (see below)
bun run type-check
bun run build
```

`test:e2e` uses your normal configuration. It runs these steps inside a fresh `sharepoint-davids-e2e-test-<timestamp>` folder:

1. Lists libraries.
2. Searches existing content and downloads a hit.
3. Creates folders.
4. Uploads a file with special characters in its name.
5. Reads the file back and compares it.
6. Checks the overwrite protection.

It then deletes the folder to the recycle bin, even when a step fails.

| Option | Default | Purpose |
| --- | --- | --- |
| `SHAREPOINT_E2E_PATH` | `Verwaltung/Documents` | Writable `<site>/<library>` to test in |
| `SHAREPOINT_E2E_QUERY` | `Rüge` | A term that already matches indexed files |
| `--wait-index` | off | Also waits up to 10 min until search finds the uploaded file |

## License

[MIT](LICENSE)
