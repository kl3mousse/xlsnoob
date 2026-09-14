import { Favorite, FavoriteGroup } from "./favorite";

export function normalizeFavoriteUrl(value: string): string {
  const trimmed = value.trim();
  try {
    const url = new URL(trimmed);
    url.protocol = url.protocol.toLowerCase();
    url.hostname = url.hostname.toLowerCase();
    url.hash = "";
    if ((url.protocol === "https:" && url.port === "443") || (url.protocol === "http:" && url.port === "80")) {
      url.port = "";
    }
    if (url.pathname.length > 1) url.pathname = url.pathname.replace(/\/+$/, "");
    return url.toString();
  } catch {
    return trimmed;
  }
}

export function deriveFavoriteName(value: string): string {
  let filename: string | undefined;
  try {
    const url = new URL(value);
    filename = decodeURIComponent(url.pathname.split("/").filter(Boolean).pop() || "");
    if (!filename) return url.hostname || "Untitled workbook";
  } catch {
    filename = value.trim().split(/[\\/]/).filter(Boolean).pop();
  }
  return filename
    ? filename.replace(/\.(xlsx|xlsm|xlsb|xls)$/i, "").replace(/_+/g, " ").trim() || "Untitled workbook"
    : "Untitled workbook";
}

export function deriveFavoriteLocation(value: string): string {
  try {
    const url = new URL(value);
    const hostname = url.hostname.replace(/^www\./i, "");
    if (hostname.toLowerCase().endsWith("-my.sharepoint.com") || /^\/personal\//i.test(url.pathname)) {
      return "Personal OneDrive";
    }
    const parts = url.pathname.split("/").filter(Boolean);
    const siteMarker = parts.findIndex((part) => /^(sites|teams)$/i.test(part));
    if (siteMarker >= 0 && parts[siteMarker + 1]) {
      return decodeURIComponent(parts[siteMarker + 1]).replace(/_+/g, " ").trim();
    }
    return hostname || "Web location";
  } catch {
    return "Saved location";
  }
}

export function filterFavorites(favorites: readonly Favorite[], query: string, groups: readonly FavoriteGroup[] = []): Favorite[] {
  const normalizedQuery = query.trim().toLocaleLowerCase();
  if (!normalizedQuery) return [...favorites];
  const groupById = new Map(groups.map((group) => [group.id, group.name]));
  return favorites.filter((favorite) => [
    favorite.name,
    favorite.comment || "",
    favorite.url,
    groupById.get(favorite.groupId || "") || "",
  ].some((value) => value.toLocaleLowerCase().includes(normalizedQuery)));
}
