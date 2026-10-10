import * as v from "valibot"

/** Runtime validation for CLI positionals and flags (parsed via sharepointArgumentParserCreate). */
export const sharepointCliArgumentSchemas = {
  remotePath: v.pipe(
    v.string(),
    v.trim(),
    v.nonEmpty("path must not be empty"),
    v.check((path) => !path.split("/").includes(".."), '".." is not allowed'),
  ),
  localPath: v.pipe(v.string(), v.nonEmpty("local path must not be empty")),
  query: v.pipe(v.string(), v.trim(), v.nonEmpty("search query must not be empty")),
  limit: v.pipe(
    v.string(),
    v.transform(Number),
    v.integer("limit must be an integer"),
    v.minValue(1, "limit must be at least 1"),
    v.maxValue(500, "limit must be at most 500"),
  ),
}
