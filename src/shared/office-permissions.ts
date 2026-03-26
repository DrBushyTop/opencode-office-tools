import { isReadOnlyOfficeTool } from "./office-tool-registry";

export interface OfficePermissionRequest {
  id: string;
  sessionID: string;
  permission: string;
  patterns: string[];
  metadata: Record<string, unknown>;
  always: string[];
}

export function toolName(request: OfficePermissionRequest) {
  return String(request.metadata.tool || request.patterns[0] || "")
}

export function canAutoApprove(request: OfficePermissionRequest) {
  if (request.permission === "doom_loop") return false
  return isReadOnlyOfficeTool(toolName(request))
}
