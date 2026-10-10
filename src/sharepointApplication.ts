import { buildApplication, buildCommand, buildRouteMap, help, version } from "@stricli/core"
import { packageVersion } from "./packageVersion.js"
import { sharepointArgumentParserCreate } from "./sharepointArgumentParserCreate.js"
import { sharepointCliArgumentSchemas as schemas } from "./sharepointCliArgumentSchemas.js"
import { sharepointCommands } from "./sharepointCommands.js"

const configFlag = {
  kind: "parsed",
  parse: sharepointArgumentParserCreate(schemas.localPath),
  optional: true,
  brief: "Config file (default: $SHAREPOINT_CONFIG or ~/.config/sharepoint/config.json)",
  placeholder: "FILE",
} as const
const jsonFlag = { kind: "boolean", default: false, brief: "Print JSON", withNegated: false } as const
const remotePath = (brief: string) =>
  ({ parse: sharepointArgumentParserCreate(schemas.remotePath), brief, placeholder: "PATH" }) as const
const remotePathOptional = (brief: string) => ({ ...remotePath(brief), optional: true }) as const

const ls = buildCommand({
  func: sharepointCommands.ls,
  parameters: {
    flags: {
      config: configFlag,
      long: { kind: "boolean", default: false, brief: "Show size and modified time", withNegated: false },
      json: jsonFlag,
    },
    aliases: { l: "long" },
    positional: { kind: "tuple", parameters: [remotePathOptional("Site, library or folder (empty: all sites)")] },
  },
  docs: { brief: "List sites, libraries or folder contents" },
})

const get = buildCommand({
  func: sharepointCommands.get,
  parameters: {
    flags: { config: configFlag },
    positional: {
      kind: "tuple",
      parameters: [
        remotePath("File to download"),
        {
          parse: sharepointArgumentParserCreate(schemas.localPath),
          brief: "Local target, '-' for stdout (default: file name)",
          placeholder: "TARGET",
          optional: true,
        },
      ],
    },
  },
  docs: { brief: "Download a file" },
})

const put = buildCommand({
  func: sharepointCommands.put,
  parameters: {
    flags: {
      config: configFlag,
      overwrite: { kind: "boolean", default: false, brief: "Replace an existing file", withNegated: false },
    },
    positional: {
      kind: "tuple",
      parameters: [
        { parse: sharepointArgumentParserCreate(schemas.localPath), brief: "Local file", placeholder: "FILE" },
        remotePath("Target folder ending in '/' (keeps the name) or target file"),
      ],
    },
  },
  docs: { brief: "Upload a file (refuses to replace existing files without --overwrite)" },
})

const mkdir = buildCommand({
  func: sharepointCommands.mkdir,
  parameters: {
    flags: { config: configFlag },
    positional: { kind: "tuple", parameters: [remotePath("Folder to create (parents included)")] },
  },
  docs: { brief: "Create a folder and missing parents" },
})

const rm = buildCommand({
  func: sharepointCommands.rm,
  parameters: {
    flags: {
      config: configFlag,
      recursive: { kind: "boolean", default: false, brief: "Allow deleting folders", withNegated: false },
    },
    aliases: { r: "recursive" },
    positional: { kind: "tuple", parameters: [remotePath("File or folder")] },
  },
  docs: { brief: "Move a file or folder to the site recycle bin" },
})

const search = buildCommand({
  func: sharepointCommands.search,
  parameters: {
    flags: {
      config: configFlag,
      limit: {
        kind: "parsed",
        parse: sharepointArgumentParserCreate(schemas.limit),
        default: "25",
        brief: "Maximum hits (1-500)",
        placeholder: "COUNT",
      },
      json: jsonFlag,
    },
    positional: {
      kind: "tuple",
      parameters: [
        {
          parse: sharepointArgumentParserCreate(schemas.query),
          brief: 'Words, "phrase", filetype:pdf, …',
          placeholder: "QUERY",
        },
      ],
    },
  },
  docs: { brief: "Full-text search across files (KQL)" },
})

/** The `sharepoint` CLI. Paths are `<site>/<library>/<path>` (or `<library>/<path>` with a configured site). */
export const sharepointApplication = buildApplication(
  buildRouteMap({
    routes: { ls, get, put, mkdir, rm, search },
    docs: {
      brief: "SharePoint files via Microsoft Graph. Paths: <site>/<library>/<path>",
    },
  }),
  { name: "sharepoint", scanner: { caseStyle: "allow-kebab-for-camel" } },
  {
    help: help({
      brief: "Print help information and exit",
      formatting: { caseStyle: "convert-camel-to-kebab", onlyRequiredInUsageLine: false, useAliasInUsageLine: false },
    }),
    version: version({ brief: "Print version information and exit", info: { currentVersion: packageVersion } }),
  },
)
