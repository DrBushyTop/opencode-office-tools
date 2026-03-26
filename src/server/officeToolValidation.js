const officeToolRegistry = require('../shared/office-tool-registry.json')

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value)
}

function validateSchema(value, schema, path = 'args') {
  if (!schema) return

  if (Array.isArray(schema.anyOf)) {
    const errors = []
    for (const item of schema.anyOf) {
      try {
        validateSchema(value, item, path)
        return
      } catch (error) {
        errors.push(error)
      }
    }
    throw new Error(`Invalid ${path}`)
  }

  if (Array.isArray(schema.enum) && !schema.enum.includes(value)) {
    throw new Error(`Invalid ${path}: expected one of ${schema.enum.join(', ')}`)
  }

  switch (schema.type) {
    case 'object': {
      if (!isPlainObject(value)) {
        throw new Error(`Invalid ${path}: expected object`)
      }

      const properties = schema.properties || {}
      const required = Array.isArray(schema.required) ? schema.required : []
      for (const name of required) {
        if (value[name] === undefined) {
          throw new Error(`Missing required ${path}.${name}`)
        }
      }

      for (const key of Object.keys(value)) {
        if (!Object.prototype.hasOwnProperty.call(properties, key)) {
          throw new Error(`Unexpected ${path}.${key}`)
        }
        validateSchema(value[key], properties[key], `${path}.${key}`)
      }
      return
    }
    case 'array': {
      if (!Array.isArray(value)) {
        throw new Error(`Invalid ${path}: expected array`)
      }
      for (let i = 0; i < value.length; i += 1) {
        validateSchema(value[i], schema.items, `${path}[${i}]`)
      }
      return
    }
    case 'string':
      if (typeof value !== 'string') throw new Error(`Invalid ${path}: expected string`)
      return
    case 'number':
      if (typeof value !== 'number' || !Number.isFinite(value)) throw new Error(`Invalid ${path}: expected number`)
      return
    case 'boolean':
      if (typeof value !== 'boolean') throw new Error(`Invalid ${path}: expected boolean`)
      return
    case undefined:
      return
    default:
      throw new Error(`Unsupported schema type for ${path}`)
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
  validateSchema(args === undefined ? {} : args, entry.parameters, 'args')
  return entry
}

module.exports = {
  validateOfficeToolCall,
  validateSchema,
}
