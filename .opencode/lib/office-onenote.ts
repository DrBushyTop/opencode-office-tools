import { tool } from "@opencode-ai/plugin"
import { hostTool } from "./office-tool"

export function onenote(name: string, description: string, args: ReturnType<typeof tool.schema.object> | Record<string, unknown>) {
  return hostTool("onenote", name, description, args)
}
