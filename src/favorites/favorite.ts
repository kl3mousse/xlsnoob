export interface FavoriteGroup {
  id: string;
  name: string;
  order: number;
  collapsed: boolean;
}

export interface Favorite {
  id: string;
  name: string;
  comment?: string;
  url: string;
  addedAt: string;
  updatedAt?: string;
  groupId?: string;
  order?: number;
  pinned?: boolean;
  [key: string]: unknown;
}

export interface FavoritesData {
  version: 2;
  groups: FavoriteGroup[];
  favorites: Favorite[];
}

export const UNGROUPED_GROUP_NAME = "Ungrouped";

export function getUngroupedGroupId(): string {
  return "ungrouped";
}

export function createDefaultGroup(): FavoriteGroup {
  return {
    id: getUngroupedGroupId(),
    name: UNGROUPED_GROUP_NAME,
    order: 0,
    collapsed: false,
  };
}

export function normalizeFavoriteGroupName(value: string): string {
  const trimmed = value.trim();
  return trimmed || UNGROUPED_GROUP_NAME;
}

export function isFavorite(value: unknown): value is Favorite {
  if (!value || typeof value !== "object") return false;
  const favorite = value as Record<string, unknown>;
  return typeof favorite.id === "string"
    && typeof favorite.name === "string"
    && typeof favorite.url === "string"
    && typeof favorite.addedAt === "string"
    && (favorite.comment === undefined || typeof favorite.comment === "string")
    && (favorite.updatedAt === undefined || typeof favorite.updatedAt === "string")
    && (favorite.groupId === undefined || typeof favorite.groupId === "string")
    && (favorite.order === undefined || typeof favorite.order === "number")
    && (favorite.pinned === undefined || typeof favorite.pinned === "boolean");
}

export function isFavoriteGroup(value: unknown): value is FavoriteGroup {
  if (!value || typeof value !== "object") return false;
  const group = value as Record<string, unknown>;
  return typeof group.id === "string"
    && typeof group.name === "string"
    && typeof group.order === "number"
    && typeof group.collapsed === "boolean";
}

export function isFavoritesData(value: unknown): value is FavoritesData {
  if (!value || typeof value !== "object") return false;
  const data = value as Record<string, unknown>;
  return data.version === 2
    && Array.isArray(data.groups)
    && Array.isArray(data.favorites);
}
