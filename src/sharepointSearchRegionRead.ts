/**
 * App-only Microsoft Search requires a region and its 400 error names the valid ones
 * ("Requested region  not found. Only valid regions are EMEA."). Returns the first named region, if any.
 */
export function sharepointSearchRegionRead(errorMessage: string): string | undefined {
  const match = /valid regions? (?:are|is):?\s*([A-Z]{2,4}(?:\s*,\s*[A-Z]{2,4})*)/i.exec(errorMessage)
  return match?.[1]?.split(",")[0]?.trim().toUpperCase()
}
