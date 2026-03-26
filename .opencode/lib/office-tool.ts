import { tool } from "@opencode-ai/plugin"
import fs from "node:fs"
import os from "node:os"
import path from "node:path"

const url = process.env.OPENCODE_OFFICE_BRIDGE_URL || "http://127.0.0.1:52391/api/office-tools/execute"

function resolveBridgeTokenPath() {
  const port = (() => {
    try {
      return String(new URL(url).port || "52391")
    } catch {
      return "52391"
    }
  })()

  return path.join(os.tmpdir(), "opencode-office-bridge", os.userInfo().username, `${port}.token`)
}

function resolveBridgeToken() {
  const tokenPath = resolveBridgeTokenPath()
  if (fs.existsSync(tokenPath)) {
    return fs.readFileSync(tokenPath, "utf8").trim()
  }

  if (process.env.OPENCODE_OFFICE_BRIDGE_TOKEN) {
    return process.env.OPENCODE_OFFICE_BRIDGE_TOKEN
  }

  return ""
}

export async function execute(host: string, name: string, args: Record<string, unknown>) {
  const bridgeToken = resolveBridgeToken()
  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      ...(bridgeToken ? { "x-office-bridge-token": bridgeToken } : {}),
    },
    body: JSON.stringify({ host, toolName: name, args }),
  })

  if (!response.ok) {
    throw new Error((await response.text()) || `Office tool failed: ${response.status}`)
  }

  const result = await response.json()
  if (typeof result.result === "string") return result.result
  return JSON.stringify(result.result, null, 2)
}

export function hostTool<Args extends typeof tool.schema.shape>(host: string, name: string, description: string, args: Args) {
  return tool({
    description,
    args,
    async execute(input) {
      return execute(host, name, input as Record<string, unknown>)
    },
  })
}
