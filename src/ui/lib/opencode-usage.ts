import { z } from "zod";
import type { OpencodeMessage } from "./opencode-schemas";

export const sessionUsageSchema = z.object({
  total: z.number().nonnegative(),
  providerID: z.string().optional(),
  modelID: z.string().optional(),
});

export type SessionUsage = z.infer<typeof sessionUsageSchema>;

const compact = new Intl.NumberFormat("en", {
  notation: "compact",
  maximumFractionDigits: 1,
});

function total(message: OpencodeMessage) {
  const tokens = message.info?.tokens;
  if (!tokens) return 0;
  const value = tokens.input + tokens.output + tokens.reasoning + tokens.cache.read + tokens.cache.write;
  if (!Number.isFinite(value) || value <= 0) return 0;
  return value;
}

export function getSessionUsage(info: OpencodeMessage["info"]) {
  const value = total({ info });
  if (value <= 0 || !info) return null;
  return sessionUsageSchema.parse({
    total: value,
    providerID: info.providerID,
    modelID: info.modelID,
  });
}

export function getLatestSessionUsage(messages: OpencodeMessage[]) {
  for (let index = messages.length - 1; index >= 0; index -= 1) {
    const message = messages[index];
    if (message.info?.role !== "assistant") continue;
    const usage = getSessionUsage(message.info);
    if (!usage) continue;
    return usage;
  }
  return null;
}

export function formatTokenUsage(
  usage: SessionUsage | null | undefined,
  models: Array<{ providerID: string; modelID: string; limitContext?: number }>,
) {
  if (!usage || usage.total <= 0) return "";
  const count = compact.format(usage.total);
  const limit = models.find((item) => item.providerID === usage.providerID && item.modelID === usage.modelID)?.limitContext;
  if (!limit || !Number.isFinite(limit) || limit <= 0) return count;
  return `${count} (${Math.round((usage.total / limit) * 100)}%)`;
}
