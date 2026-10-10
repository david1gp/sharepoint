# Goal
Create `@adaptive-ds/sharepoint`, a Bun/TypeScript library and `sharepoint` CLI for SharePoint files via Microsoft Graph. It replaces the bundled CLI in `leo_internal/dev-servers/philipp-server/sharepoint` and is published as the public MIT repository `david1gp/sharepoint`.

# Decisions
- Use Microsoft Graph via `@pnp/graph` and `@pnp/nodejs` with app-only client-secret auth. SharePoint REST (`@pnp/sp`) would need a certificate for app-only access.
- Config sources: `--config`, then `SHAREPOINT_CONFIG`, then `~/.config/sharepoint/config.json`, with `SHAREPOINT_*` variables overriding single fields. Only `tenantId`, `clientId` and `clientSecret` are required. `hostname` is discovered from the root site and `searchRegion` from Microsoft Search's own error message. An optional `site` pins all paths to one site.
- `XDG_CONFIG_HOME` is deliberately ignored: per-scope coding agents set it and would otherwise lose the shared config.
- Paths are `<site>/<library>/<path>`, the same layout as the rclone `sharepoint:` combine remote used in dev-servers. rclone config generation stays in dev-servers, because the package does not need rclone.
- Safety: uploads never replace a file without `--overwrite`, `rm` needs `-r` for folders and refuses whole libraries, and deletes go to the site recycle bin.
- Use `@stricli/core` for the CLI, valibot for config and argument validation, and `Result` from `@adaptive-ds/result`. Follow the code-style skill: one export per file and subject-first names.
- Unit and CLI tests run against an in-memory Graph fake. `SHAREPOINT_GRAPH_URL` is accepted only for `localhost`.
- `test:e2e` runs live against the tenant, outside `bun test`. It works inside `sharepoint-davids-e2e-test-<timestamp>` under `Verwaltung/Documents` and always cleans up. Search is checked against already-indexed content, because indexing new files takes minutes; `--wait-index` is opt-in.
- npm OIDC/trusted publishing, the first manual publish and releases are done by David.

# Tasks
1. Complete: Copy the template files from telegram-send and set up package metadata, scripts and the workspace.
2. Complete: Library (config, client, list, download, upload, mkdir, delete, search) and the Stricli CLI.
3. Complete: Tests (config, paths, operations and CLI against the Graph fake), plus the e2e script verified live.
4. Complete: README, build and package contents checked; Node and Bun runtimes checked.
5. Complete: Git init, initial commit, public GitHub repo, push (no release tag).
6. Pending, after the first publish: migrate philipp-server (global install, `~/.config/sharepoint/config.json`, remove `SHAREPOINT_*` from env.conf, AGENTS.md, run e2e on the server).
