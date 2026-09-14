import { Favorite, FavoriteGroup, FavoritesData, createDefaultGroup, getUngroupedGroupId, isFavorite, isFavoritesData } from "./favorite";
import { FavoritesStore } from "./favorites-store";

const favoritesStorageKey = "xlsnoob.favorites.v2";
const legacyFavoritesStorageKey = "xlsnoob.favorites.v1";

interface StorageLike {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}

function normalizeFavoriteList(favorites: unknown[]): Favorite[] {
  return favorites.filter(isFavorite).map((favorite, index) => ({
    ...favorite,
    groupId: favorite.groupId && favorite.groupId !== "" ? favorite.groupId : getUngroupedGroupId(),
    order: typeof favorite.order === "number" ? favorite.order : index,
    pinned: favorite.pinned === true ? true : undefined,
  }));
}

function normalizeGroupList(groups: unknown[]): FavoriteGroup[] {
  const normalized = groups.filter((group): group is FavoriteGroup => {
    if (!group || typeof group !== "object") return false;
    const candidate = group as Record<string, unknown>;
    return typeof candidate.id === "string" && typeof candidate.name === "string" && typeof candidate.order === "number" && typeof candidate.collapsed === "boolean";
  });
  const withDefaultGroup = normalized.some((group) => group.id === getUngroupedGroupId())
    ? normalized
    : [createDefaultGroup(), ...normalized];
  return withDefaultGroup
    .map((group, index) => ({
      id: group.id,
      name: group.name.trim() || "Ungrouped",
      order: typeof group.order === "number" ? group.order : index,
      collapsed: Boolean(group.collapsed),
    }))
    .sort((left, right) => left.order - right.order)
    .map((group, index) => ({ ...group, order: index }));
}

function normalizeFavoritesForGroups(favorites: unknown[], groups: FavoriteGroup[]): Favorite[] {
  const ungroupedId = getUngroupedGroupId();
  return normalizeFavoriteList(favorites).map((favorite) => ({
    ...favorite,
    groupId: favorite.groupId && groups.some((group) => group.id === favorite.groupId) ? favorite.groupId : ungroupedId,
    order: typeof favorite.order === "number" ? favorite.order : 0,
  }));
}

function createEmptyData(): FavoritesData {
  return {
    version: 2,
    groups: [createDefaultGroup()],
    favorites: [],
  };
}

function migrateLegacyData(value: unknown): FavoritesData {
  const favorites = Array.isArray(value) ? normalizeFavoriteList(value) : [];
  const defaultGroups = [createDefaultGroup()];
  const normalizedFavorites = favorites.map((favorite, index) => ({
    ...favorite,
    groupId: getUngroupedGroupId(),
    order: index,
  }));
  return {
    version: 2,
    groups: defaultGroups,
    favorites: normalizedFavorites,
  };
}

export class LocalStorageFavoritesStore implements FavoritesStore {
  constructor(private readonly storage: StorageLike = localStorage) {}

  async list(): Promise<Favorite[]> {
    return (await this.listData()).favorites;
  }

  async listData(): Promise<FavoritesData> {
    const value = this.storage.getItem(favoritesStorageKey);
    if (value) {
      try {
        const parsed: unknown = JSON.parse(value);
        if (isFavoritesData(parsed)) {
          const groups = normalizeGroupList(parsed.groups);
          return {
            version: 2,
            groups,
            favorites: normalizeFavoritesForGroups(parsed.favorites, groups),
          };
        }
      } catch {
        // fall through to migration
      }
    }

    const legacyValue = this.storage.getItem(legacyFavoritesStorageKey);
    if (legacyValue) {
      try {
        const parsed: unknown = JSON.parse(legacyValue);
        const migrated = migrateLegacyData(parsed);
        this.saveData(migrated);
        return migrated;
      } catch {
        // fall through
      }
    }

    const empty = createEmptyData();
    this.saveData(empty);
    return empty;
  }

  async saveData(data: FavoritesData): Promise<void> {
    const groups = normalizeGroupList(data.groups.length ? data.groups : [createDefaultGroup()]);
    const favorites = normalizeFavoritesForGroups(data.favorites, groups);
    const normalizedData: FavoritesData = {
      version: 2,
      groups,
      favorites,
    };
    this.storage.setItem(favoritesStorageKey, JSON.stringify(normalizedData));
  }

  async add(favorite: Favorite): Promise<void> {
    const data = await this.listData();
    data.favorites.push({
      ...favorite,
      groupId: favorite.groupId ?? getUngroupedGroupId(),
      order: typeof favorite.order === "number" ? favorite.order : data.favorites.length,
    });
    await this.saveData(data);
  }

  async update(favorite: Favorite): Promise<void> {
    const data = await this.listData();
    const index = data.favorites.findIndex((item) => item.id === favorite.id);
    if (index < 0) throw new Error("Favorite no longer exists.");
    data.favorites[index] = { ...favorite, groupId: favorite.groupId ?? getUngroupedGroupId(), order: typeof favorite.order === "number" ? favorite.order : 0 };
    await this.saveData(data);
  }

  async remove(id: string): Promise<void> {
    const data = await this.listData();
    data.favorites = data.favorites.filter((favorite) => favorite.id !== id);
    await this.saveData(data);
  }

  async reorder(ids: string[]): Promise<void> {
    const data = await this.listData();
    if (ids.length !== data.favorites.length || new Set(ids).size !== data.favorites.length) {
      throw new Error("Favorites could not be reordered because the list changed.");
    }
    const byId = new Map(data.favorites.map((favorite) => [favorite.id, favorite]));
    const reordered = ids.map((id) => byId.get(id));
    if (reordered.some((favorite) => !favorite)) {
      throw new Error("Favorites could not be reordered because the list changed.");
    }
    reordered.forEach((favorite, index) => {
      if (favorite) favorite.order = index;
    });
    await this.saveData(data);
  }
}
