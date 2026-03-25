import { tool } from "@opencode-ai/plugin"
import { hostTool } from "./office-tool"

export function powerpoint<Args extends typeof tool.schema.shape>(name: string, description: string, args: Args) {
  return hostTool("powerpoint", name, description, args)
}
