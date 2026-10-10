import * as v from "valibot"

const guidSchema = v.pipe(v.string(), v.trim(), v.uuid("must be a GUID (Entra → App registrations → Overview)"))

/**
 * Only tenantId/clientId/clientSecret are required. hostname is discovered from the tenant root site,
 * searchRegion from Microsoft Search itself; site pins every path to one site (paths become `<library>/<path>`).
 */
export const sharepointConfigSchema = v.strictObject({
  tenantId: guidSchema,
  clientId: guidSchema,
  clientSecret: v.pipe(
    v.string(),
    v.trim(),
    v.minLength(20, "must be the client secret *value*, not the secret ID"),
    v.regex(/^\S+$/, "must not contain whitespace"),
  ),
  hostname: v.optional(
    v.pipe(
      v.string(),
      v.trim(),
      v.toLowerCase(),
      v.regex(/^[a-z0-9-]+\.sharepoint\.(com|us|de|cn)$/, "must look like <tenant>.sharepoint.com"),
    ),
  ),
  site: v.optional(v.pipe(v.string(), v.trim(), v.nonEmpty(), v.regex(/^[^/]+$/, "must be a site name, e.g. Kunden"))),
  searchRegion: v.optional(v.pipe(v.string(), v.trim(), v.toUpperCase(), v.regex(/^[A-Z]{2,4}$/))),
})

export type SharepointConfig = v.InferOutput<typeof sharepointConfigSchema>
