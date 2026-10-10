const hints: Record<number, string> = {
  401: "check tenantId/clientId/clientSecret and that the secret has not expired",
  403: "the app lacks access: grant Graph Sites.ReadWrite.All (Application) with admin consent, or this site via Sites.Selected",
}

/**
 * PnP throws errors whose message ends in "::> {json}"; reduce them to the HTTP status plus Graph's own message.
 * MSAL errors (bad tenant/secret) carry no status and keep their message.
 */
export function sharepointErrorRead(error: unknown): { status: number; message: string } {
  const status = typeof error === "object" && error !== null && "status" in error ? Number(error.status) || 0 : 0
  const raw = error instanceof Error ? error.message : String(error)
  const body = raw.split("::> ")[1]
  let message = raw
  if (body) {
    try {
      const parsed: unknown = JSON.parse(body)
      const graphError = typeof parsed === "object" && parsed !== null && "error" in parsed ? parsed.error : undefined
      if (typeof graphError === "object" && graphError !== null && "message" in graphError)
        message = String(graphError.message)
    } catch {
      message = raw
    }
  }
  const hint = hints[status]
  const withStatus = status ? `${status} ${message}` : message
  return { status, message: hint ? `${withStatus} (${hint})` : withStatus }
}
