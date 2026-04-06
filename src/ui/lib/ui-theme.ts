import { useSyncExternalStore } from "react";

export type ThemeMode = "light" | "dark";
export type ThemePreference = ThemeMode | "system";

type ThemeTokens = {
  background: string;
  text: string;
  border: string;
  accent: string;
};

type ThemeSurface = ThemeTokens & {
  page: string;
  panel: string;
  panelStrong: string;
  soft: string;
  softHover: string;
  subtle: string;
  input: string;
  muted: string;
  faint: string;
  userBubble: string;
  userBorder: string;
  success: string;
  successSoft: string;
  danger: string;
  dangerSoft: string;
  shadow: string;
  icon: string;
};

type ThemeDefinition = {
  id: string;
  label: string;
  light: ThemeSurface;
  dark: ThemeSurface;
};

export const defaultThemeId = "oc-2";

const themes: ThemeDefinition[] = [
  {
    id: "oc-2",
    label: "OC-2",
    light: {
      background: "#F8F8F8",
      text: "#171717",
      border: "#DBDBDB",
      accent: "#034cff",
      page: "#F8F8F8",
      panel: "#FCFCFC",
      panelStrong: "#FFFFFF",
      soft: "rgba(0, 0, 0, 0.031)",
      softHover: "rgba(0, 0, 0, 0.059)",
      subtle: "rgba(0, 0, 0, 0.051)",
      input: "#FCFCFC",
      muted: "#6F6F6F",
      faint: "#8F8F8F",
      userBubble: "rgba(0, 0, 0, 0.031)",
      userBorder: "#E5E5E5",
      success: "#1D7A43",
      successSoft: "rgba(29, 122, 67, 0.10)",
      danger: "#ED4831",
      dangerSoft: "#FFF2F0",
      shadow: "0 0 0 1px rgba(0, 0, 0, 0.07), 0 36px 80px rgba(0, 0, 0, 0.03)",
      icon: "#8F8F8F",
    },
    dark: {
      background: "#1C1C1C",
      text: "#EDEDED",
      border: "#282828",
      accent: "#034cff",
      page: "#161616",
      panel: "#1C1C1C",
      panelStrong: "#202020",
      soft: "rgba(255, 255, 255, 0.06)",
      softHover: "rgba(255, 255, 255, 0.10)",
      subtle: "rgba(255, 255, 255, 0.08)",
      input: "#202020",
      muted: "#B8B8B8",
      faint: "#989898",
      userBubble: "rgba(255, 255, 255, 0.04)",
      userBorder: "rgba(255, 255, 255, 0.10)",
      success: "#8EE6A3",
      successSoft: "rgba(29, 122, 67, 0.22)",
      danger: "#FE806A",
      dangerSoft: "rgba(252, 83, 58, 0.16)",
      shadow: "0 0 0 1px rgba(255, 255, 255, 0.06), 0 36px 80px rgba(0, 0, 0, 0.28)",
      icon: "#A8A8A8",
    },
  },
  {
    id: "catppuccin",
    label: "Catppuccin",
    light: {
      background: "#f5e0dc",
      text: "#4c4f69",
      border: "#d9c3c0",
      accent: "#d20f39",
      page: "#f2deda",
      panel: "#f5e0dc",
      panelStrong: "#f8e8e4",
      soft: "rgba(76, 79, 105, 0.08)",
      softHover: "rgba(76, 79, 105, 0.12)",
      subtle: "rgba(76, 79, 105, 0.10)",
      input: "#f8e8e4",
      muted: "#6c7086",
      faint: "#7c7f93",
      userBubble: "rgba(114, 135, 253, 0.10)",
      userBorder: "rgba(114, 135, 253, 0.22)",
      success: "#40a02b",
      successSoft: "rgba(64, 160, 43, 0.12)",
      danger: "#d20f39",
      dangerSoft: "rgba(210, 15, 57, 0.12)",
      shadow: "0 0 0 1px rgba(76, 79, 105, 0.10), 0 36px 80px rgba(76, 79, 105, 0.10)",
      icon: "#6c7086",
    },
    dark: {
      background: "#1e1e2e",
      text: "#cdd6f4",
      border: "#45475a",
      accent: "#f38ba8",
      page: "#181825",
      panel: "#1e1e2e",
      panelStrong: "#24273a",
      soft: "rgba(205, 214, 244, 0.08)",
      softHover: "rgba(205, 214, 244, 0.12)",
      subtle: "rgba(205, 214, 244, 0.10)",
      input: "#24273a",
      muted: "#a6adc8",
      faint: "#9399b2",
      userBubble: "rgba(180, 190, 254, 0.12)",
      userBorder: "rgba(180, 190, 254, 0.22)",
      success: "#a6d189",
      successSoft: "rgba(166, 209, 137, 0.16)",
      danger: "#f38ba8",
      dangerSoft: "rgba(243, 139, 168, 0.16)",
      shadow: "0 0 0 1px rgba(205, 214, 244, 0.08), 0 36px 80px rgba(0, 0, 0, 0.32)",
      icon: "#a6adc8",
    },
  },
  {
    id: "catppuccin-frappe",
    label: "Catppuccin Frappe",
    light: {
      background: "#303446",
      text: "#c6d0f5",
      border: "#b5bfe2",
      accent: "#f4b8e4",
      page: "#232634",
      panel: "#303446",
      panelStrong: "#414559",
      soft: "rgba(198, 208, 245, 0.08)",
      softHover: "rgba(198, 208, 245, 0.12)",
      subtle: "rgba(198, 208, 245, 0.10)",
      input: "#414559",
      muted: "#b5bfe2",
      faint: "#949cb8",
      userBubble: "rgba(141, 164, 226, 0.16)",
      userBorder: "rgba(141, 164, 226, 0.24)",
      success: "#a6d189",
      successSoft: "rgba(166, 209, 137, 0.16)",
      danger: "#e78284",
      dangerSoft: "rgba(231, 130, 132, 0.16)",
      shadow: "0 0 0 1px rgba(181, 191, 226, 0.10), 0 36px 80px rgba(0, 0, 0, 0.36)",
      icon: "#b5bfe2",
    },
    dark: {
      background: "#303446",
      text: "#c6d0f5",
      border: "#b5bfe2",
      accent: "#f4b8e4",
      page: "#232634",
      panel: "#303446",
      panelStrong: "#414559",
      soft: "rgba(198, 208, 245, 0.08)",
      softHover: "rgba(198, 208, 245, 0.12)",
      subtle: "rgba(198, 208, 245, 0.10)",
      input: "#414559",
      muted: "#b5bfe2",
      faint: "#949cb8",
      userBubble: "rgba(141, 164, 226, 0.16)",
      userBorder: "rgba(141, 164, 226, 0.24)",
      success: "#a6d189",
      successSoft: "rgba(166, 209, 137, 0.16)",
      danger: "#e78284",
      dangerSoft: "rgba(231, 130, 132, 0.16)",
      shadow: "0 0 0 1px rgba(181, 191, 226, 0.10), 0 36px 80px rgba(0, 0, 0, 0.36)",
      icon: "#b5bfe2",
    },
  },
  {
    id: "catppuccin-macchiato",
    label: "Catppuccin Macchiato",
    light: {
      background: "#24273a",
      text: "#cad3f5",
      border: "#b8c0e0",
      accent: "#f5bde6",
      page: "#1e2030",
      panel: "#24273a",
      panelStrong: "#363a4f",
      soft: "rgba(202, 211, 245, 0.08)",
      softHover: "rgba(202, 211, 245, 0.12)",
      subtle: "rgba(202, 211, 245, 0.10)",
      input: "#363a4f",
      muted: "#b8c0e0",
      faint: "#939ab7",
      userBubble: "rgba(138, 173, 244, 0.16)",
      userBorder: "rgba(138, 173, 244, 0.24)",
      success: "#a6da95",
      successSoft: "rgba(166, 218, 149, 0.16)",
      danger: "#ed8796",
      dangerSoft: "rgba(237, 135, 150, 0.16)",
      shadow: "0 0 0 1px rgba(184, 192, 224, 0.10), 0 36px 80px rgba(0, 0, 0, 0.36)",
      icon: "#b8c0e0",
    },
    dark: {
      background: "#24273a",
      text: "#cad3f5",
      border: "#b8c0e0",
      accent: "#f5bde6",
      page: "#1e2030",
      panel: "#24273a",
      panelStrong: "#363a4f",
      soft: "rgba(202, 211, 245, 0.08)",
      softHover: "rgba(202, 211, 245, 0.12)",
      subtle: "rgba(202, 211, 245, 0.10)",
      input: "#363a4f",
      muted: "#b8c0e0",
      faint: "#939ab7",
      userBubble: "rgba(138, 173, 244, 0.16)",
      userBorder: "rgba(138, 173, 244, 0.24)",
      success: "#a6da95",
      successSoft: "rgba(166, 218, 149, 0.16)",
      danger: "#ed8796",
      dangerSoft: "rgba(237, 135, 150, 0.16)",
      shadow: "0 0 0 1px rgba(184, 192, 224, 0.10), 0 36px 80px rgba(0, 0, 0, 0.36)",
      icon: "#b8c0e0",
    },
  },
];

export const themeOptions = themes.map((theme) => ({
  id: theme.id,
  label: theme.label,
  isDefault: theme.id === defaultThemeId,
}));

function readColorSchemeQuery() {
  if (typeof window === "undefined" || typeof window.matchMedia !== "function") {
    return null;
  }

  return window.matchMedia("(prefers-color-scheme: dark)");
}

export function readSystemThemeMode(): ThemeMode {
  return readColorSchemeQuery()?.matches ? "dark" : "light";
}

export function subscribeToSystemThemeMode(onStoreChange: () => void) {
  const mediaQuery = readColorSchemeQuery();
  if (!mediaQuery) {
    return () => undefined;
  }

  const handleChange = () => {
    onStoreChange();
  };

  mediaQuery.addEventListener("change", handleChange);
  return () => {
    mediaQuery.removeEventListener("change", handleChange);
  };
}

export function resolveThemeModePreference(preference: ThemePreference, systemMode: ThemeMode): ThemeMode {
  return preference === "system" ? systemMode : preference;
}

export function useThemeMode(preference: ThemePreference): ThemeMode {
  const systemMode = useSyncExternalStore<ThemeMode>(subscribeToSystemThemeMode, readSystemThemeMode, () => "light");
  return resolveThemeModePreference(preference, systemMode);
}

function pickTheme(themeId: string) {
  return themes.find((theme) => theme.id === themeId) || themes[0]!;
}

export function resolveThemeTokens(themeId: string) {
  const theme = pickTheme(themeId);
  return {
    id: theme.id,
    label: theme.label,
    light: {
      background: theme.light.background,
      text: theme.light.text,
      border: theme.light.border,
      accent: theme.light.accent,
    },
    dark: {
      background: theme.dark.background,
      text: theme.dark.text,
      border: theme.dark.border,
      accent: theme.dark.accent,
    },
  };
}

export function getThemeCssVars(themeId: string, mode: ThemeMode) {
  const theme = pickTheme(themeId);
  const values = mode === "dark" ? theme.dark : theme.light;

  return {
    "--background-base": values.page,
    "--background-weak": values.background,
    "--background-strong": values.panel,
    "--background-stronger": values.panelStrong,
    "--surface-base": values.soft,
    "--surface-base-hover": values.softHover,
    "--surface-weak": values.subtle,
    "--surface-raised-base": values.soft,
    "--surface-raised-strong": values.panel,
    "--surface-inset-base": values.subtle,
    "--surface-interactive-base": values.userBubble,
    "--surface-interactive-hover": values.softHover,
    "--input-base": values.input,
    "--text-base": values.muted,
    "--text-weak": values.faint,
    "--text-weaker": values.faint,
    "--text-strong": values.text,
    "--text-interactive-base": values.accent,
    "--text-on-interactive-base": "#FCFCFC",
    "--text-on-critical-base": values.danger,
    "--border-base": values.border,
    "--border-weak-base": values.border,
    "--border-strong-base": values.border,
    "--border-interactive-base": values.accent,
    "--border-critical-base": values.danger,
    "--icon-base": values.icon,
    "--icon-strong-base": values.text,
    "--button-primary-base": values.text,
    "--shadow-lg-border-base": values.shadow,
    "--oc-page": values.page,
    "--oc-bg": values.panel,
    "--oc-bg-strong": values.panelStrong,
    "--oc-bg-elevated": values.panelStrong,
    "--oc-bg-soft": values.soft,
    "--oc-bg-soft-hover": values.softHover,
    "--oc-border": values.border,
    "--oc-border-strong": values.border,
    "--oc-text": values.text,
    "--oc-text-muted": values.muted,
    "--oc-text-faint": values.faint,
    "--oc-accent": values.accent,
    "--oc-accent-strong": values.accent,
    "--oc-accent-bg": values.userBubble,
    "--oc-thinking-bg": values.soft,
    "--oc-thinking-border": values.subtle,
    "--oc-thinking-label": "var(--oc-warning)",
    "--oc-thinking-title": "color-mix(in srgb, var(--oc-warning) 82%, var(--oc-text) 18%)",
    "--oc-shadow": values.shadow,
    "--oc-danger-bg": values.dangerSoft,
    "--oc-danger-border": values.danger,
    "--oc-danger-text": values.danger,
    "--oc-danger": values.danger,
    "--oc-warning": "#d4a72c",
    "--oc-success": values.success,
    "--oc-success-soft": values.successSoft,
    "--oc-user-bubble": values.userBubble,
    "--oc-user-border": values.userBorder,
  } satisfies Record<string, string>;
}
