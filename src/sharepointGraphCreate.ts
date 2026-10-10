import { graphfi } from "@pnp/graph"
import "@pnp/graph/sites/index.js"
import "@pnp/graph/files/index.js"
import "@pnp/graph/search/index.js"
import { GraphDefault, MSAL } from "@pnp/nodejs"
import type { SharepointConfig } from "./sharepointConfigSchema.js"

/**
 * App-only Microsoft Graph client (client credentials). SharePoint's own REST API rejects client secrets for
 * app-only tokens, which is why everything here goes through Graph. `baseUrl` without auth is for test fakes.
 */
export function sharepointGraphCreate(
  config: Pick<SharepointConfig, "tenantId" | "clientId" | "clientSecret">,
  options: Readonly<{ baseUrl?: string }> = {},
) {
  if (options.baseUrl) return graphfi().using(GraphDefault({ baseUrl: options.baseUrl }))
  return graphfi().using(
    GraphDefault(),
    MSAL(
      {
        auth: {
          authority: `https://login.microsoftonline.com/${config.tenantId}`,
          clientId: config.clientId,
          clientSecret: config.clientSecret,
        },
      },
      ["https://graph.microsoft.com/.default"],
    ),
  )
}

export type SharepointGraph = ReturnType<typeof sharepointGraphCreate>
