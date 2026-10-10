import * as v from "valibot"

/** Stricli parsers must throw on invalid input; valibot's message becomes the CLI error. */
export function sharepointArgumentParserCreate<const T extends v.GenericSchema<string, unknown>>(schema: T) {
  return (input: string): v.InferOutput<T> => {
    const parsed = v.safeParse(schema, input)
    if (!parsed.success) throw new Error(v.summarize(parsed.issues))
    return parsed.output
  }
}
