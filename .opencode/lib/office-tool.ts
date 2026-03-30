import { tool } from "@opencode-ai/plugin"
import fs from "node:fs"
import os from "node:os"
import path from "node:path"

const url = process.env.OPENCODE_OFFICE_BRIDGE_URL || "http://127.0.0.1:52391/api/office-tools/execute"

type Binary = {
  data: string
  mimeType: string
  type: string
  description?: string
}

type Result = {
  textResultForLlm?: string
  binaryResultsForLlm?: Binary[]
}

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

function ext(mime: string) {
  const kind = mime.split("/")[1] || "bin"
  if (kind === "jpeg") return "jpg"
  if (kind === "svg+xml") return "svg"
  return kind.replace(/[^a-zA-Z0-9]+/g, "-") || "bin"
}

function outdir() {
  const dir = path.join(os.tmpdir(), "opencode-office-tool-output")
  fs.mkdirSync(dir, { recursive: true })
  return dir
}

function persist(name: string, items: Binary[]) {
  return items.map((item, i) => {
    const file = path.join(outdir(), `${name}-${Date.now()}-${crypto.randomUUID()}-${i + 1}.${ext(item.mimeType)}`)
    fs.writeFileSync(file, Buffer.from(item.data, "base64"))
    return file
  })
}

function readHint(result: Result, files: string[]) {
  if (!files.length) return undefined
  const image = result.binaryResultsForLlm?.every((item) => item.mimeType.startsWith("image/"))
  if (image) {
    return `Call the Read tool immediately on ${files.length === 1 ? "the image path above" : "each image path above"} to inspect the returned image content.`
  }
  return "Use the Read tool on a file path above to inspect the returned content."
}

function format(name: string, result: Result) {
  const files = result.binaryResultsForLlm?.length ? persist(name, result.binaryResultsForLlm) : []
  const text = typeof result.textResultForLlm === "string" ? result.textResultForLlm : ""
  if (!files.length) return text || JSON.stringify(result, null, 2)
  return [
    text || `Saved ${files.length} file${files.length === 1 ? "" : "s"}.`,
    "Saved tool files:",
    ...files.map((file) => `- ${file}`),
    readHint(result, files),
  ].filter(Boolean).join("\n")
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
  if (result.result && typeof result.result === "object") {
    return format(name, result.result as Result)
  }
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
