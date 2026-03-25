import { tool } from "@opencode-ai/plugin"
import { hostTool } from "./office-tool"

export function word<Args extends typeof tool.schema.shape>(name: string, description: string, args: Args) {
  return hostTool("word", name, description, args)
}
