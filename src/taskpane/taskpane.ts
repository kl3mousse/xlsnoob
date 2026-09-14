/// <reference types="office-js" />

import { openContainingFolder, openInfo } from "../commands/file";
import { initializeFavorites } from "../favorites/favorites-ui";

type ThemeName = "light" | "dark";

const themeStorageKey = "xlsnoob.theme";
const sharedThemeTokens: Record<string, string> = {
  "--brand-ribbon-orange": "#217346",
  "--text-on-accent": "#ffffff",
  "--shadow-subtle": "0 1px 2px rgba(0, 0, 0, 0.1)",
  "--radius-sm": "4px",
  "--radius-md": "7px",
  "--radius-lg": "8px",
  "--space-xs": "4px",
  "--space-sm": "8px",
  "--space-md": "12px",
  "--space-lg": "16px",
};
const themeTokens: Record<ThemeName, Record<string, string>> = {
  light: {
    "--bg": "#f5f5f6",
    "--surface": "#ffffff",
    "--surface-hover": "#fafafa",
    "--surface-active": "#f1f1f1",
    "--text": "#242424",
    "--text-muted": "#6f6f6f",
    "--text-subtle": "#929292",
    "--border": "#e4e4e4",
    "--border-strong": "#d2d2d2",
    "--accent": "#217346",
    "--accent-hover": "#185c37",
    "--accent-soft": "#e4f1ea",
    "--accent-text": "#17472c",
    "--danger": "#b42318",
    "--danger-soft": "#fff1f0",
    "--success": "#498205",
    "--warning": "#b77a00",
    "--shadow-menu": "0 8px 24px rgba(0, 0, 0, 0.14)",
  },
  dark: {
    "--bg": "#1f1f1f",
    "--surface": "#292929",
    "--surface-hover": "#333333",
    "--surface-active": "#3a3a3a",
    "--text": "#f3f3f3",
    "--text-muted": "#b8b8b8",
    "--text-subtle": "#8f8f8f",
    "--border": "#3d3d3d",
    "--border-strong": "#505050",
    "--accent": "#217346",
    "--accent-hover": "#35a165",
    "--accent-soft": "#163322",
    "--accent-text": "#bce8cd",
    "--danger": "#ffb4ab",
    "--danger-soft": "#4a201d",
    "--success": "#a9d476",
    "--warning": "#ffd66b",
    "--shadow-menu": "0 8px 24px rgba(0, 0, 0, 0.4)",
  },
};

function getSystemTheme(): ThemeName {
  return typeof window.matchMedia === "function" && window.matchMedia("(prefers-color-scheme: dark)").matches
    ? "dark"
    : "light";
}

function getSavedTheme(): ThemeName {
  try {
    const savedTheme = localStorage.getItem(themeStorageKey);
    if (savedTheme === "light" || savedTheme === "dark") return savedTheme;
  } catch {
    // Use the system setting when storage is unavailable.
  }
  return getSystemTheme();
}

function setTheme(theme: ThemeName): void {
  const root = document.documentElement;
  root.dataset.theme = theme;
  Object.entries({ ...sharedThemeTokens, ...themeTokens[theme] }).forEach(([name, value]) => {
    root.style.setProperty(name, value);
  });
  const toggle = document.getElementById("themeToggle") as HTMLButtonElement | null;
  if (toggle) {
    const nextTheme = theme === "dark" ? "light" : "dark";
    toggle.setAttribute("aria-label", `Use ${nextTheme} theme`);
    toggle.title = `Use ${nextTheme} theme`;
  }
}

function initializeTheme(): void {
  setTheme(getSavedTheme());
  document.getElementById("themeToggle")?.addEventListener("click", () => {
    const nextTheme: ThemeName = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
    setTheme(nextTheme);
    try {
      localStorage.setItem(themeStorageKey, nextTheme);
    } catch {
      // The current session still reflects the selected theme.
    }
  });
}

function disableAutomaticTaskPane(): void {
  Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", false);
  Office.context.document.settings.saveAsync();
}

initializeTheme();

Office.onReady(() => {
  disableAutomaticTaskPane();
  initializeFavorites();
});

Object.assign(globalThis, {
  openContainingFolder,
  openInfo,
});
