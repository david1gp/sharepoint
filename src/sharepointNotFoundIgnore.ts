import { sharepointErrorRead } from "./sharepointErrorRead.js"

/** For existence probes: a 404 becomes `undefined`, every other error is rethrown to the caller's catch. */
export async function sharepointNotFoundIgnore<T>(request: Promise<T>): Promise<T | undefined> {
  try {
    return await request
  } catch (error) {
    if (sharepointErrorRead(error).status === 404) return undefined
    throw error
  }
}
