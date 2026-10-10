import { expect, test } from "bun:test"
import { sharepointConfigLoad } from "../src/sharepointConfigLoad.js"

const ids = {
  tenantId: "00000000-0000-0000-0000-000000000001",
  clientId: "00000000-0000-0000-0000-000000000002",
  clientSecret: "abc~defghijklmnopqrstuvwxyz",
}
const missingFile = async () => {
  throw Object.assign(new Error("missing"), { code: "ENOENT" })
}
const fileWith =
  (files: Record<string, unknown>) =>
  async (path: string): Promise<string> => {
    if (!(path in files)) return missingFile()
    return JSON.stringify(files[path])
  }

test("config file alone is enough: only tenantId, clientId and clientSecret are required", async () => {
  const result = await sharepointConfigLoad({
    environment: {},
    home: "/home/u",
    readFile: fileWith({ "/home/u/.config/sharepoint/config.json": ids }),
  })
  expect(result.success && result.data).toEqual(ids)
})

test("SHAREPOINT_* variables override individual config file fields", async () => {
  const result = await sharepointConfigLoad({
    environment: { SHAREPOINT_SITE: "Kunden", SHAREPOINT_CLIENT_SECRET: "env-secret-value-0123456789" },
    home: "/home/u",
    readFile: fileWith({ "/home/u/.config/sharepoint/config.json": { ...ids, site: "Verwaltung" } }),
  })
  expect(result.success && result.data.site).toBe("Kunden")
  expect(result.success && result.data.clientSecret).toBe("env-secret-value-0123456789")
})

test("variables alone work when no config file exists", async () => {
  const result = await sharepointConfigLoad({
    environment: {
      SHAREPOINT_TENANT_ID: ids.tenantId,
      SHAREPOINT_CLIENT_ID: ids.clientId,
      SHAREPOINT_CLIENT_SECRET: ids.clientSecret,
    },
    readFile: missingFile,
  })
  expect(result.success).toBe(true)
})

test("SHAREPOINT_CONFIG and --config select another file, and an explicit file must exist", async () => {
  const readFile = fileWith({ "/etc/sp.json": ids })
  expect((await sharepointConfigLoad({ environment: { SHAREPOINT_CONFIG: "/etc/sp.json" }, readFile })).success).toBe(
    true,
  )
  expect((await sharepointConfigLoad({ configPath: "/etc/sp.json", environment: {}, readFile })).success).toBe(true)
  const missing = await sharepointConfigLoad({ configPath: "/nope.json", environment: {}, readFile })
  expect(!missing.success && missing.errorMessage).toContain("/nope.json")
})

test("missing configuration explains where to put it", async () => {
  const result = await sharepointConfigLoad({ environment: {}, home: "/home/u", readFile: missingFile })
  expect(!result.success && result.errorMessage).toContain("/home/u/.config/sharepoint/config.json")
})

test("invalid values, unknown keys and a secret ID instead of the secret value are rejected", async () => {
  const load = (config: unknown) =>
    sharepointConfigLoad({
      environment: {},
      home: "/h",
      readFile: fileWith({ "/h/.config/sharepoint/config.json": config }),
    })
  expect((await load({ ...ids, tenantId: "contoso" })).success).toBe(false)
  expect((await load({ ...ids, clientSecret: "c236b8c6-ac25" })).success).toBe(false)
  expect((await load({ ...ids, hostname: "evil.example.com" })).success).toBe(false)
  expect((await load({ ...ids, sitee: "typo" })).success).toBe(false)
})
