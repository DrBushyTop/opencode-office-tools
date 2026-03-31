const { z, ZodError } = require('zod')
const officeToolRegistry = require('../shared/office-tool-registry.json')

const schemaCache = new WeakMap()

function isBlankString(value) {
  return typeof value === 'string' && value.trim() === ''
}

function hasNonBlankString(value) {
  return typeof value === 'string' && value.trim() !== ''
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
        if (typeof schema.minItems === 'number') {
          builtSchema = builtSchema.min(schema.minItems)
        }
        if (typeof schema.maxItems === 'number') {
          builtSchema = builtSchema.max(schema.maxItems)
        }
        break
      case 'string':
        builtSchema = z.string()
        if (typeof schema.minLength === 'number') {
          builtSchema = builtSchema.min(schema.minLength)
        }
        if (typeof schema.maxLength === 'number') {
          builtSchema = builtSchema.max(schema.maxLength)
        }
        if (typeof schema.pattern === 'string') {
          builtSchema = builtSchema.regex(new RegExp(schema.pattern))
        }
        break
      case 'number':
        builtSchema = z.number().finite()
        if (typeof schema.minimum === 'number') {
          builtSchema = builtSchema.min(schema.minimum)
        }
        if (typeof schema.maximum === 'number') {
          builtSchema = builtSchema.max(schema.maximum)
        }
        if (typeof schema.exclusiveMinimum === 'number') {
          builtSchema = builtSchema.gt(schema.exclusiveMinimum)
        }
        if (typeof schema.exclusiveMaximum === 'number') {
          builtSchema = builtSchema.lt(schema.exclusiveMaximum)
        }
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

  if ((toolName === 'read_slide_text' || toolName === 'edit_slide_text') && isBlankString(normalizedArgs.ref)) {
    throw new Error('Invalid args.ref: cannot be empty')
  }

  if (toolName === 'edit_slide_xml' && Array.isArray(normalizedArgs.replacements)) {
    if (normalizedArgs.replacements.length === 0) {
      throw new Error('Invalid args.replacements: provide at least one replacement')
    }

    const invalidBlankRefIndex = normalizedArgs.replacements.findIndex((replacement) => replacement && isBlankString(replacement.ref))
    if (invalidBlankRefIndex >= 0) {
      throw new Error(`Invalid args.replacements[${invalidBlankRefIndex}].ref: cannot be empty`)
    }
  }

  if (toolName === 'edit_slide_chart' && isBlankString(normalizedArgs.ref)) {
    throw new Error('Invalid args.ref: cannot be empty')
  }

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

  if (toolName === 'edit_slide_xml') {
    if (!Array.isArray(normalizedArgs.replacements) || normalizedArgs.replacements.length === 0) {
      throw new Error('Invalid args.replacements: provide at least one replacement')
    }

    const invalidReplacementIndex = normalizedArgs.replacements.findIndex((replacement) => !replacement || isBlankString(replacement.ref))
    if (invalidReplacementIndex >= 0) {
      throw new Error(`Invalid args.replacements[${invalidReplacementIndex}].ref: cannot be empty`)
    }
  }

  if (toolName === 'edit_slide_chart') {
    if ((normalizedArgs.action === 'update' || normalizedArgs.action === 'delete') && !hasNonBlankString(normalizedArgs.ref)) {
      throw new Error(`Missing required args.ref for ${normalizedArgs.action} actions`)
    }

    if (normalizedArgs.action === 'create' || normalizedArgs.action === 'update') {
      if (!hasNonBlankString(normalizedArgs.chartType)) {
        throw new Error(`Missing required args.chartType for ${normalizedArgs.action} actions`)
      }
      if (!Array.isArray(normalizedArgs.series) || normalizedArgs.series.length === 0) {
        throw new Error(`Missing required args.series for ${normalizedArgs.action} actions`)
      }
    }
  }

  if (toolName === 'manage_slide_media') {
    if ((normalizedArgs.action === 'insertImage' || normalizedArgs.action === 'replaceImage') && !hasNonBlankString(normalizedArgs.imageUrl)) {
      throw new Error('Missing required args.imageUrl for insertImage and replaceImage actions')
    }
    if ((normalizedArgs.action === 'replaceImage' || normalizedArgs.action === 'deleteImage') && normalizedArgs.shapeId === undefined) {
      throw new Error('Missing required args.shapeId for replaceImage and deleteImage actions')
    }
  }

  if (toolName === 'manage_slide_table') {
    if ((normalizedArgs.action === 'create' || normalizedArgs.action === 'update')) {
      if (!Array.isArray(normalizedArgs.values) || normalizedArgs.values.length === 0 || !Array.isArray(normalizedArgs.values[0]) || normalizedArgs.values[0].length === 0) {
        throw new Error('Invalid args.values: must be a non-empty 2D array for create and update actions')
      }
    }
    if ((normalizedArgs.action === 'update' || normalizedArgs.action === 'delete') && normalizedArgs.shapeId === undefined) {
      throw new Error('Missing required args.shapeId for update and delete actions')
    }
  }

  if (toolName === 'add_slide_animation') {
    if (Array.isArray(normalizedArgs.shapeId) && normalizedArgs.shapeId.length === 0) {
      throw new Error('Invalid args.shapeId: array must not be empty')
    }
    if (normalizedArgs.durationMs !== undefined && normalizedArgs.durationMs < 0) {
      throw new Error('Invalid args.durationMs: must be a non-negative number')
    }
    if (normalizedArgs.delayMs !== undefined && normalizedArgs.delayMs < 0) {
      throw new Error('Invalid args.delayMs: must be a non-negative number')
    }
    if (normalizedArgs.repeatCount !== undefined && normalizedArgs.repeatCount < 0) {
      throw new Error('Invalid args.repeatCount: must be a non-negative number')
    }
    if (normalizedArgs.type === 'motionPath' && !hasNonBlankString(normalizedArgs.path)) {
      throw new Error('Missing required args.path for motionPath animations')
    }
    if (normalizedArgs.type === 'scale' && normalizedArgs.scaleXPercent === undefined && normalizedArgs.scaleYPercent === undefined) {
      throw new Error('Missing required args.scaleXPercent or args.scaleYPercent for scale animations')
    }
    if (normalizedArgs.type === 'rotate' && normalizedArgs.angleDegrees === undefined) {
      throw new Error('Missing required args.angleDegrees for rotate animations')
    }
    if ((normalizedArgs.type === 'complementaryColor' || normalizedArgs.type === 'changeFillColor' || normalizedArgs.type === 'changeLineColor') && !hasNonBlankString(normalizedArgs.toColor)) {
      throw new Error('Missing required args.toColor for emphasis color animations')
    }
  }

  if (toolName === 'set_slide_transition') {
    if (Array.isArray(normalizedArgs.slideIndex) && normalizedArgs.slideIndex.length === 0) {
      throw new Error('Invalid args.slideIndex: array must not be empty')
    }
    if (normalizedArgs.advanceAfterMs !== undefined && normalizedArgs.advanceAfterMs < 0) {
      throw new Error('Invalid args.advanceAfterMs: must be a non-negative number')
    }
    if (normalizedArgs.durationMs !== undefined && normalizedArgs.durationMs < 0) {
      throw new Error('Invalid args.durationMs: must be a non-negative number')
    }
    if (['push', 'wipe', 'cover', 'pull'].includes(normalizedArgs.effect) && normalizedArgs.direction !== undefined && !['left', 'right', 'up', 'down'].includes(normalizedArgs.direction)) {
      throw new Error(`Invalid args.direction: ${normalizedArgs.effect} transitions only support left, right, up, or down`)
    }
    if (normalizedArgs.effect === 'randomBar' && normalizedArgs.direction !== undefined && !['horizontal', 'vertical'].includes(normalizedArgs.direction)) {
      throw new Error('Invalid args.direction: randomBar transitions only support horizontal or vertical')
    }
    if (normalizedArgs.effect === 'split') {
      if (normalizedArgs.direction !== undefined && !['in', 'out'].includes(normalizedArgs.direction)) {
        throw new Error('Invalid args.direction: split transitions only support in or out')
      }
      if (normalizedArgs.orientation !== undefined && !['horizontal', 'vertical'].includes(normalizedArgs.orientation)) {
        throw new Error('Invalid args.orientation: split transitions only support horizontal or vertical')
      }
    }
  }

  if (toolName === 'edit_slide_master') {
    const hasThemeColors = normalizedArgs.themeColors
      && typeof normalizedArgs.themeColors === 'object'
      && Object.keys(normalizedArgs.themeColors).length > 0
    const hasDecorativeShapes = Array.isArray(normalizedArgs.decorativeShapes) && normalizedArgs.decorativeShapes.length > 0

    if (!hasThemeColors && !hasDecorativeShapes) {
      throw new Error('Provide args.themeColors or args.decorativeShapes')
    }
  }

  if (toolName === 'create_slide_from_layout' && Array.isArray(normalizedArgs.bindings)) {
    const invalidBindingIndex = normalizedArgs.bindings.findIndex((binding) => !binding
      || (!hasNonBlankString(binding.placeholderType) && !hasNonBlankString(binding.placeholderName)))
    if (invalidBindingIndex >= 0) {
      throw new Error(`Invalid args.bindings[${invalidBindingIndex}]: each binding must include placeholderType or placeholderName`)
    }

    const invalidTableIndex = normalizedArgs.bindings.findIndex((binding) => {
      if (!binding || !Array.isArray(binding.tableData) || binding.tableData.length === 0) return false
      if (!Array.isArray(binding.tableData[0]) || binding.tableData[0].length === 0) return true
      const expectedWidth = binding.tableData[0].length
      return binding.tableData.some((row) => !Array.isArray(row) || row.length !== expectedWidth || row.length === 0)
    })
    if (invalidTableIndex >= 0) {
      throw new Error(`Invalid args.bindings[${invalidTableIndex}].tableData: must be a non-empty rectangular 2D array`)
    }
  }

  if (toolName === 'add_slide_from_code' && isBlankString(normalizedArgs.code)) {
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
