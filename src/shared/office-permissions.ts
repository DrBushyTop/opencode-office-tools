import { isReadOnlyOfficeTool } from "./office-tool-registry";
import type { OfficePermissionRequest } from "./office-metadata";
export { officePermissionRequestSchema } from "./office-metadata";
export type { OfficePermissionRequest } from "./office-metadata";

const OUTPUT_DIR = /(^|[\\/])opencode-office-tool-output([\\/]|$)/;

export function toolName(request: OfficePermissionRequest) {
  return String(request.metadata.tool || request.patterns[0] || "")
}

export function permissionTarget(request: OfficePermissionRequest) {
  if (request.permission === "task") {
    return String(request.metadata.subagent_type || request.metadata.description || request.patterns[0] || "subagent")
  }

  if (request.permission === "edit") {
    return String(request.metadata.filepath || request.patterns[0] || "file")
  }

  if (request.permission === "read") {
    return String(request.metadata.filepath || request.patterns[0] || "file")
  }

  return String(request.metadata.tool || request.patterns[0] || "")
}

export function permissionKind(request: OfficePermissionRequest) {
  if (request.permission === "doom_loop") return "danger"
  if (request.permission === "task") return "subagent"
  if (["read", "glob", "grep", "list", "todoread"].includes(request.permission)) return "read"
  if (["edit", "write", "todowrite"].includes(request.permission)) return "write"
  if (request.permission === "bash") return "shell"
  if (request.permission === "mcp") return "mcp"
  if (request.permission === "tool") {
    return isReadOnlyOfficeTool(toolName(request)) ? "read" : "write"
  }
  return "generic"
}

export function canAutoApprove(request: OfficePermissionRequest) {
  if (request.permission === "doom_loop") return false
  if (request.permission === "tool") {
    return isReadOnlyOfficeTool(toolName(request))
  }

  if (request.permission === "read" || request.permission === "external_directory") {
    return request.patterns.some((item) => OUTPUT_DIR.test(item))
  }

  return false
}
