import { parseRgbColor, isDarkRgb } from "./utilities.js";

export function getTeamsThemeHint(): string | null {
  const classNames = [document.documentElement.className, document.body?.className || ""].join(" ");

  if (/\btheme-(dark|black|contrast)(?:[a-z0-9-]*)\b|\bdark\b|\bcontrast\b/i.test(classNames)) {
    return "dark";
  }

  if (/\btheme-(default|light)(?:[a-z0-9-]*)\b|\blight\b/i.test(classNames)) {
    return "light";
  }

  return null;
}

export function getPageThemeFallback(): string {
  const probes = [
    document.querySelector('[data-tid="title-bar"]'),
    document.querySelector('[data-tid="titlebar-end-slot"]'),
    document.body,
    document.documentElement
  ].filter(Boolean) as Element[];

  for (const probe of probes) {
    const backgroundColor = parseRgbColor(window.getComputedStyle(probe).backgroundColor);
    const dark = isDarkRgb(backgroundColor);
    if (dark !== null) {
      return dark ? "dark" : "light";
    }
  }

  return window.matchMedia?.("(prefers-color-scheme: dark)")?.matches ? "dark" : "light";
}

export function detectTheme(): string {
  return getTeamsThemeHint() || getPageThemeFallback();
}
