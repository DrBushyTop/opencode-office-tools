const { z, ZodError } = require('zod')
const officeToolRegistry = require('../shared/office-tool-registry.json')

const schemaCache = new WeakMap()

function isBlankString(value) {
  return typeof value === 'string' && value.trim() === ''
}

function joinIssuePath(basePath, segments = []) {
  return segments.reduce((current, segment) => (
    typeof segment === 'number' ? `${current}[${segment}]` : `${current}.${segment}`
  ), basePath)
}

function formatZodError(error, basePath = 'args') {
  if (!(error instanceof ZodError)) return error

  const issue = error.issues[0]
  if (!issue) return new Error(`Invalid ${basePath}`)

  const path = joinIssuePath(basePath, issue.path)

  if (issue.code === 'unrecognized_keys' && issue.keys.length > 0) {
    return new Error(`Unexpected ${joinIssuePath(basePath, [issue.keys[0]])}`)
  }

  if (issue.code === 'invalid_type') {
    if (issue.input === undefined) {
      return new Error(`Missing required ${path}`)
    }
    return new Error(`Invalid ${path}: expected ${issue.expected}`)
  }

  if (issue.code === 'invalid_value' && Array.isArray(issue.values)) {
    return new Error(`Invalid ${path}: expected one of ${issue.values.join(', ')}`)
  }

  if (issue.code === 'invalid_union') {
    return new Error(`Invalid ${path}`)
  }

  return new Error(issue.message ? `Invalid ${path}: ${issue.message}` : `Invalid ${path}`)
}

function buildZodSchema(schema, path = 'args') {
  if (!schema) return z.any()
  if (schemaCache.has(schema)) return schemaCache.get(schema)

  let builtSchema

  if (Array.isArray(schema.anyOf)) {
    const options = schema.anyOf.map((item) => buildZodSchema(item, path))
    builtSchema = options.length === 1 ? options[0] : z.union(options)
  } else if (Array.isArray(schema.enum)) {
    if (schema.enum.length > 0 && schema.enum.every((value) => typeof value === 'string')) {
      builtSchema = z.enum(schema.enum)
    } else {
      const options = schema.enum.map((value) => z.literal(value))
      builtSchema = options.length === 1 ? options[0] : z.union(options)
    }
  } else {
    switch (schema.type) {
      case 'object': {
        const properties = schema.properties || {}
        const required = new Set(Array.isArray(schema.required) ? schema.required : [])
        const shape = {}

        for (const [key, value] of Object.entries(properties)) {
          const propertySchema = buildZodSchema(value, `${path}.${key}`)
          shape[key] = required.has(key) ? propertySchema : propertySchema.optional()
        }

        builtSchema = z.object(shape).strict()
        break
      }
      case 'array':
        builtSchema = z.array(buildZodSchema(schema.items, `${path}[]`))
        break
      case 'string':
        builtSchema = z.string()
        break
      case 'number':
        builtSchema = z.number().finite()
        break
      case 'boolean':
        builtSchema = z.boolean()
        break
      case undefined:
        builtSchema = z.any()
        break
      default:
        throw new Error(`Unsupported schema type for ${path}`)
    }
  }

  schemaCache.set(schema, builtSchema)
  return builtSchema
}

function validateSchema(value, schema, path = 'args') {
  try {
    buildZodSchema(schema, path).parse(value)
  } catch (error) {
    throw formatZodError(error, path)
  }
}

function validateOfficeToolCall(host, toolName, args) {
  if (!Object.prototype.hasOwnProperty.call(officeToolRegistry, toolName)) {
    throw new Error(`Unknown Office tool '${toolName}'`)
  }

  const entry = officeToolRegistry[toolName]
  if (!entry) {
    throw new Error(`Unknown Office tool '${toolName}'`)
  }
  if (!entry.hosts.includes(host)) {
    throw new Error(`Tool '${toolName}' is not available for host '${host}'`)
  }
  const normalizedArgs = args === undefined ? {} : args
  validateSchema(normalizedArgs, entry.parameters, 'args')

  if (toolName === 'manage_range'
    && normalizedArgs.action === 'filter'
    && (normalizedArgs.filterOperation === undefined || normalizedArgs.filterOperation === 'apply')) {
    const hasCriteria = Boolean(
      normalizedArgs.filterOn
      || normalizedArgs.criterion1
      || normalizedArgs.criterion2
      || (Array.isArray(normalizedArgs.filterValues) && normalizedArgs.filterValues.length)
      || normalizedArgs.dynamicCriteria
      || normalizedArgs.filterColor,
    )

    if (hasCriteria && normalizedArgs.columnIndex === undefined) {
      throw new Error('Missing required args.columnIndex when applying filter criteria')
    }
  }

  if (toolName === 'set_document_range') {
    const operation = normalizedArgs.operation === undefined ? 'replace' : normalizedArgs.operation
    if ((operation === 'replace' || operation === 'insert') && normalizedArgs.content === undefined) {
      throw new Error('Missing required args.content for replace or insert operations')
    }
  }

  if (toolName === 'navigate_to_page') {
    const hasPageId = typeof normalizedArgs.pageId === 'string' && normalizedArgs.pageId.trim() !== ''
    const hasClientUrl = typeof normalizedArgs.clientUrl === 'string' && normalizedArgs.clientUrl.trim() !== ''
    if (hasPageId === hasClientUrl) {
      throw new Error('Provide exactly one of args.pageId or args.clientUrl')
    }
  }

  if (toolName === 'set_note_selection' && isBlankString(normalizedArgs.content)) {
    throw new Error('Invalid args.content: cannot be empty')
  }

  if ((toolName === 'edit_slide_with_code' || toolName === 'add_slide_from_code') && isBlankString(normalizedArgs.code)) {
    throw new Error('Invalid args.code: cannot be empty')
  }

  if (toolName === 'set_page_title' && isBlankString(normalizedArgs.title)) {
    throw new Error('Invalid args.title: cannot be empty')
  }

  if (toolName === 'append_page_content' && isBlankString(normalizedArgs.html)) {
    throw new Error('Invalid args.html: cannot be empty')
  }

  return entry
}

module.exports = {
  validateOfficeToolCall,
  validateSchema,
}
