import { homedir } from "node:os"
import { join } from "node:path"

/**
 * --config flag, then SHAREPOINT_CONFIG, then ~/.config/sharepoint/config.json. XDG_CONFIG_HOME is deliberately
 * ignored: sandboxed tools (e.g. per-scope coding agents) often point it elsewhere and would lose the shared config.
 */
export function sharepointConfigPathResolve(
  options: Readonly<{ configPath?: string; environment: Readonly<Record<string, string | undefined>>; home?: string }>,
): { path: string; explicit: boolean } {
  if (options.configPath) return { path: options.configPath, explicit: true }
  const fromEnvironment = options.environment.SHAREPOINT_CONFIG?.trim()
  if (fromEnvironment) return { path: fromEnvironment, explicit: true }
  return { path: join(options.home ?? homedir(), ".config", "sharepoint", "config.json"), explicit: false }
}
