/** Graph's `root:/<path>:` addressing takes the path verbatim; encode each segment so spaces, umlauts, # and ? survive. */
export function sharepointItemPathEncode(itemPath: string): string {
  return itemPath.split("/").filter(Boolean).map(encodeURIComponent).join("/")
}
