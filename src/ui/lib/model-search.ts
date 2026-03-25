import type { ModelInfo } from "./opencode-client";

function normalize(value: string) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function score(query: string, value: string) {
  const needle = normalize(query);
  const hay = normalize(value);
  if (!needle) return 0;
  if (!hay) return Number.NEGATIVE_INFINITY;

  let pos = -1;
  let total = 0;
  let streak = 0;

  for (const char of needle) {
    const next = hay.indexOf(char, pos + 1);
    if (next === -1) return Number.NEGATIVE_INFINITY;

    const gap = pos === -1 ? next : next - pos - 1;
    streak = next === pos + 1 ? streak + 1 : 1;
    total += 12 - Math.min(gap, 10) + streak * 6;
    if (next === 0) total += 10;
    pos = next;
  }

  total -= hay.length - needle.length;
  return total;
}

function fields(model: Pick<ModelInfo, "key" | "label" | "providerID" | "modelID">) {
  return [model.label, model.key, model.providerID, model.modelID];
}

export function filterModels<T extends Pick<ModelInfo, "key" | "label" | "providerID" | "modelID">>(
  models: T[],
  query: string,
) {
  const value = query.trim();
  if (!value) return models;

  return models
    .map((model) => {
      const best = Math.max(...fields(model).map((item) => score(value, item)));
      return { model, best };
    })
    .filter((item) => item.best > Number.NEGATIVE_INFINITY)
    .sort((a, b) => b.best - a.best || a.model.label.localeCompare(b.model.label))
    .map((item) => item.model);
}
