import type { ResultErr } from "#result"

/** Stricli prints a returned Error and sets a non-zero exit code. */
export function sharepointCliError(result: ResultErr): Error {
  return new Error(`sharepoint: ${result.errorMessage}`)
}
