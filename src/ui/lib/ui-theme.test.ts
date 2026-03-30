import { describe, expect, it } from "vitest";

const load = () => {
  // @ts-expect-error TDD: helper module is intentionally not implemented yet.
  return import("./ui-theme");
};

describe("ui theme helper", () => {
  it("exposes a catalog of theme options with oc-2 as the default", async () => {
    const mod = await load();

    expect(mod.defaultThemeId).toBe("oc-2");
    expect(mod.themeOptions).toEqual(expect.arrayContaining([
      { id: "oc-2", label: "OC-2", isDefault: true },
      { id: "catppuccin", label: "Catppuccin", isDefault: false },
      { id: "catppuccin-frappe", label: "Catppuccin Frappe", isDefault: false },
      { id: "catppuccin-macchiato", label: "Catppuccin Macchiato", isDefault: false },
    ]));
  });

  it("resolves oc-2 into chat styling tokens for light and dark modes", async () => {
    const mod = await load();

    expect(mod.resolveThemeTokens("oc-2")).toEqual({
      id: "oc-2",
      label: "OC-2",
      light: {
        background: "#F8F8F8",
        text: "#171717",
        border: "#DBDBDB",
        accent: "#034cff",
      },
      dark: {
        background: "#1C1C1C",
        text: "#EDEDED",
        border: "#282828",
        accent: "#034cff",
      },
    });
  });

  it("resolves catppuccin variants into chat styling tokens", async () => {
    const mod = await load();

    expect(mod.resolveThemeTokens("catppuccin")).toMatchObject({
      id: "catppuccin",
      label: "Catppuccin",
      light: {
        background: "#f5e0dc",
        text: "#4c4f69",
        accent: "#d20f39",
      },
      dark: {
        background: "#1e1e2e",
        text: "#cdd6f4",
        accent: "#f38ba8",
      },
    });

    expect(mod.resolveThemeTokens("catppuccin-frappe")).toMatchObject({
      id: "catppuccin-frappe",
      label: "Catppuccin Frappe",
      light: {
        background: "#303446",
        text: "#c6d0f5",
        border: "#b5bfe2",
        accent: "#f4b8e4",
      },
      dark: {
        background: "#303446",
        text: "#c6d0f5",
        border: "#b5bfe2",
        accent: "#f4b8e4",
      },
    });

    expect(mod.resolveThemeTokens("catppuccin-macchiato")).toMatchObject({
      id: "catppuccin-macchiato",
      label: "Catppuccin Macchiato",
      light: {
        background: "#24273a",
        text: "#cad3f5",
        border: "#b8c0e0",
        accent: "#f5bde6",
      },
      dark: {
        background: "#24273a",
        text: "#cad3f5",
        border: "#b8c0e0",
        accent: "#f5bde6",
      },
    });

    const catppuccin = mod.resolveThemeTokens("catppuccin");
    expect(catppuccin.light.border).toEqual(expect.any(String));
    expect(catppuccin.dark.border).toEqual(expect.any(String));
  });

  it("falls back to oc-2 when the theme id is unknown", async () => {
    const mod = await load();

    expect(mod.resolveThemeTokens("missing-theme")).toEqual(mod.resolveThemeTokens("oc-2"));
  });
});
