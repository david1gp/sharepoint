import type { SharepointConfig } from "./sharepointConfigSchema.js"

const sharepointEnvironmentKeys = {
  SHAREPOINT_TENANT_ID: "tenantId",
  SHAREPOINT_CLIENT_ID: "clientId",
  SHAREPOINT_CLIENT_SECRET: "clientSecret",
  SHAREPOINT_HOSTNAME: "hostname",
  SHAREPOINT_SITE: "site",
  SHAREPOINT_SEARCH_REGION: "searchRegion",
} as const satisfies Record<string, keyof SharepointConfig>

/** `SHAREPOINT_*` variables → partial config; empty values count as unset. */
export function sharepointEnvironmentConfigRead(
  environment: Readonly<Record<string, string | undefined>>,
): Partial<Record<keyof SharepointConfig, string>> {
  const config: Partial<Record<keyof SharepointConfig, string>> = {}
  for (const [variable, key] of Object.entries(sharepointEnvironmentKeys)) {
    const value = environment[variable]?.trim()
    if (value) config[key] = value
  }
  return config
}
