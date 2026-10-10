/** decodeURIComponent that returns the input unchanged for malformed escapes (e.g. a literal "%" in a name). */
export function sharepointUrlDecode(value: string): string {
  try {
    return decodeURIComponent(value)
  } catch {
    return value
  }
}
