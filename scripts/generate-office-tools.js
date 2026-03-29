const fs = require("fs")
const path = require("path")

const registry = require("../src/shared/office-tool-registry.json")

const outputDir = path.join(__dirname, "..", ".opencode", "tools")

const hostHelpers = {
  word: { fn: "word", importPath: "../lib/office-word" },
  excel: { fn: "excel", importPath: "../lib/office-excel" },
  powerpoint: { fn: "powerpoint", importPath: "../lib/office-powerpoint" },
  onenote: { fn: "onenote", importPath: "../lib/office-onenote" },
}

function indent(text, spaces) {
  const pad = " ".repeat(spaces)
  return text.split("\n").map((line) => `${pad}${line}`).join("\n")
}

function schemaCode(schema) {
  if (schema.anyOf) {
    return `tool.schema.union([${schema.anyOf.map((item) => schemaCode(item)).join(", ")}])`
  }

  if (schema.type === "object") {
    const properties = schema.properties || {}
    const required = schema.required || []
    const entries = Object.keys(properties).map((name) => propertyCode(name, properties[name], required))
    return `tool.schema.object({\n${entries.join("\n")}\n})`
  }

  if (schema.type === "array") {
    return `tool.schema.array(${schemaCode(schema.items || { type: "string" })})`
  }

  if (schema.type === "string" && Array.isArray(schema.enum)) {
    return `tool.schema.enum([${schema.enum.map((value) => JSON.stringify(value)).join(", ")}])`
  }

  switch (schema.type) {
    case "string":
      return "tool.schema.string()"
    case "number":
      return "tool.schema.number()"
    case "boolean":
      return "tool.schema.boolean()"
    default:
      throw new Error(`Unsupported schema: ${JSON.stringify(schema)}`)
  }
}

function propertyCode(name, schema, required) {
  let code = schemaCode(schema)
  if (!required.includes(name)) {
    code += ".optional()"
  }
  if (schema.description) {
    code += `.describe(${JSON.stringify(schema.description)})`
  }
  return `  ${name}: ${code},`
}

function fileCode(name, entry) {
  if (!Array.isArray(entry.hosts) || entry.hosts.length !== 1) {
    throw new Error(`Generated wrappers require exactly one host for ${name}`)
  }

  const host = entry.hosts[0]
  const helper = hostHelpers[host]
  if (!helper) throw new Error(`Unsupported host for ${name}: ${host}`)

  const properties = entry.parameters?.properties || {}
  const required = entry.parameters?.required || []
  const propertyNames = Object.keys(properties)
  const needsToolImport = propertyNames.length > 0

  const lines = []
  if (needsToolImport) {
    lines.push('import { tool } from "@opencode-ai/plugin"')
  }
  lines.push(`import { ${helper.fn} } from "${helper.importPath}"`)
  lines.push("")

  if (propertyNames.length === 0) {
    lines.push(`export default ${helper.fn}(${JSON.stringify(name)}, ${JSON.stringify(entry.description)}, {})`)
    return `${lines.join("\n")}\n`
  }

  lines.push(`export default ${helper.fn}(${JSON.stringify(name)}, ${JSON.stringify(entry.description)}, {`)
  for (const propertyName of propertyNames) {
    lines.push(propertyCode(propertyName, properties[propertyName], required))
  }
  lines.push("})")
  return `${lines.join("\n")}\n`
}

function main() {
  fs.mkdirSync(outputDir, { recursive: true })

  const generatedNames = new Set()
  for (const [name, entry] of Object.entries(registry)) {
    const filePath = path.join(outputDir, `${name}.ts`)
    fs.writeFileSync(filePath, fileCode(name, entry))
    generatedNames.add(`${name}.ts`)
  }

  for (const fileName of fs.readdirSync(outputDir)) {
    if (fileName.endsWith(".ts") && !generatedNames.has(fileName)) {
      fs.unlinkSync(path.join(outputDir, fileName))
    }
  }
}

main()
