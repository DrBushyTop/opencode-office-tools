import { tool } from "@opencode-ai/plugin"

const url = process.env.OPENCODE_OFFICE_BRIDGE_URL || "http://127.0.0.1:52391/api/office-tools/execute"

export async function execute(name: string, args: Record<string, unknown>) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ host: "word", toolName: name, args }),
  })

  if (!response.ok) {
    throw new Error((await response.text()) || `Office tool failed: ${response.status}`)
  }

  const result = await response.json()
  if (typeof result.result === "string") return result.result
  return JSON.stringify(result.result, null, 2)
}

export function word<Args extends typeof tool.schema.shape>(name: string, description: string, args: Args) {
  return tool({
    description,
    args,
    async execute(input) {
      return execute(name, input as Record<string, unknown>)
    },
  })
}
