export interface OfficePermissionRequest {
  id: string;
  sessionID: string;
  permission: string;
  patterns: string[];
  metadata: Record<string, unknown>;
  always: string[];
}

const readOnly = new Set([
  "get_document_overview",
  "get_document_content",
  "get_document_section",
  "get_selection",
  "get_selection_text",
]);

export function toolName(request: OfficePermissionRequest) {
  return String(request.metadata.tool || request.patterns[0] || "")
}

export function canAutoApprove(request: OfficePermissionRequest) {
  if (request.permission === "doom_loop") return false
  return readOnly.has(toolName(request))
}
