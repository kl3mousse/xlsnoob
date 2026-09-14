import { Favorite, FavoriteGroup, FavoritesData, getUngroupedGroupId, normalizeFavoriteGroupName } from "./favorite";
import { FavoritesStore } from "./favorites-store";
import { deriveFavoriteName, normalizeFavoriteUrl } from "./favorite-utils";

export interface FavoriteChanges {
  name: string;
  comment?: string;
  url: string;
  groupId?: string;
  orderedAt?: number;
}

function createFavoriteId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
}

export class FavoritesService {
  constructor(
    private readonly store: FavoritesStore,
    private readonly createId: () => string = createFavoriteId,
    private readonly now: () => Date = () => new Date(),
  ) {}

  async list(): Promise<Favorite[]> {
    const data = await this.listData();
    const groups = [...data.groups].sort((left, right) => left.order - right.order);
    const orderByGroup = new Map(groups.map((group, index) => [group.id, index]));
    return [...data.favorites].sort((left, right) => {
      const leftGroup = orderByGroup.get(left.groupId ?? getUngroupedGroupId()) ?? 0;
      const rightGroup = orderByGroup.get(right.groupId ?? getUngroupedGroupId()) ?? 0;
      if (leftGroup !== rightGroup) return leftGroup - rightGroup;
      return (left.order ?? 0) - (right.order ?? 0);
    });
  }

  async listGroups(): Promise<FavoriteGroup[]> {
    return [...(await this.listData()).groups].sort((left, right) => left.order - right.order);
  }

  async listData(): Promise<FavoritesData> {
    return this.store.listData();
  }

  async findByUrl(url: string): Promise<Favorite | undefined> {
    const normalizedUrl = normalizeFavoriteUrl(url);
    return (await this.list()).find((favorite) => normalizeFavoriteUrl(favorite.url) === normalizedUrl);
  }

  private nextGroupOrder(data: FavoritesData): number {
    return data.groups.reduce((max, group) => Math.max(max, group.order), -1) + 1;
  }

  private nextFavoriteOrder(data: FavoritesData, groupId: string): number {
    const groupFavorites = data.favorites.filter((favorite) => (favorite.groupId ?? getUngroupedGroupId()) === groupId);
    return groupFavorites.reduce((max, favorite) => Math.max(max, favorite.order ?? 0), -1) + 1;
  }

  async add(url: string, name = deriveFavoriteName(url), comment?: string): Promise<Favorite> {
    const existing = await this.findByUrl(url);
    if (existing) return existing;
    const data = await this.listData();
    const groupId = getUngroupedGroupId();
    const favorite: Favorite = {
      id: this.createId(),
      name: name.trim() || deriveFavoriteName(url),
      url: url.trim(),
      addedAt: this.now().toISOString(),
      groupId,
      order: this.nextFavoriteOrder(data, groupId),
    };
    if (comment?.trim()) favorite.comment = comment.trim();
    data.favorites.push(favorite);
    await this.store.saveData(data);
    return favorite;
  }

  async addGroup(name: string): Promise<FavoriteGroup> {
    const trimmedName = normalizeFavoriteGroupName(name);
    const data = await this.listData();
    const group: FavoriteGroup = {
      id: this.createId(),
      name: trimmedName,
      order: this.nextGroupOrder(data),
      collapsed: false,
    };
    data.groups.push(group);
    await this.store.saveData(data);
    return group;
  }

  async renameGroup(id: string, name: string): Promise<FavoriteGroup> {
    const trimmedName = normalizeFavoriteGroupName(name);
    const data = await this.listData();
    const group = data.groups.find((item) => item.id === id);
    if (!group) throw new Error("Group no longer exists.");
    group.name = trimmedName;
    await this.store.saveData(data);
    return group;
  }

  async updateGroupCollapsed(id: string, collapsed: boolean): Promise<FavoriteGroup> {
    const data = await this.listData();
    const group = data.groups.find((item) => item.id === id);
    if (!group) throw new Error("Group no longer exists.");
    group.collapsed = collapsed;
    await this.store.saveData(data);
    return group;
  }

  async moveGroup(id: string, offset: -1 | 1): Promise<void> {
    const data = await this.listData();
    const groups = [...data.groups].sort((left, right) => left.order - right.order);
    const index = groups.findIndex((group) => group.id === id);
    if (index < 0) return;
    const destination = index + offset;
    if (destination < 0 || destination >= groups.length) return;
    [groups[index], groups[destination]] = [groups[destination], groups[index]];
    groups.forEach((group, newIndex) => {
      group.order = newIndex;
    });
    data.groups = groups;
    await this.store.saveData(data);
  }

  async deleteGroup(id: string): Promise<void> {
    if (id === getUngroupedGroupId()) return;
    const data = await this.listData();
    const group = data.groups.find((item) => item.id === id);
    if (!group) return;
    data.groups = data.groups.filter((item) => item.id !== id);
    let nextUngroupedOrder = this.nextFavoriteOrder(data, getUngroupedGroupId());
    data.favorites.forEach((favorite) => {
      if ((favorite.groupId ?? getUngroupedGroupId()) === id) {
        favorite.groupId = getUngroupedGroupId();
        favorite.order = nextUngroupedOrder++;
      }
    });
    data.groups.forEach((item, index) => {
      item.order = index;
    });
    await this.store.saveData(data);
  }

  async update(id: string, changes: FavoriteChanges): Promise<Favorite> {
    const data = await this.listData();
    const favorite = data.favorites.find((item) => item.id === id);
    if (!favorite) throw new Error("Favorite no longer exists.");
    const duplicate = data.favorites.find((item) => item.id !== id
      && normalizeFavoriteUrl(item.url) === normalizeFavoriteUrl(changes.url));
    if (duplicate) throw new Error("That workbook is already in Favorites.");
    const updated: Favorite = {
      ...favorite,
      name: changes.name.trim() || deriveFavoriteName(changes.url),
      url: changes.url.trim(),
      updatedAt: this.now().toISOString(),
      groupId: changes.groupId ?? favorite.groupId ?? getUngroupedGroupId(),
      order: typeof favorite.order === "number" ? favorite.order : 0,
    };
    if (changes.comment?.trim()) updated.comment = changes.comment.trim();
    else delete updated.comment;
    const index = data.favorites.findIndex((item) => item.id === id);
    data.favorites[index] = updated;
    await this.store.saveData(data);
    return updated;
  }

  remove(id: string): Promise<void> {
    return this.store.remove(id);
  }

  reorder(ids: string[]): Promise<void> {
    return this.store.reorder(ids);
  }

  async reorderGroupFavorites(groupId: string, ids: string[]): Promise<void> {
    const data = await this.listData();
    const favorites = data.favorites.filter((favorite) => (favorite.groupId ?? getUngroupedGroupId()) === groupId);
    if (ids.length !== favorites.length || new Set(ids).size !== favorites.length) {
      throw new Error("Favorites could not be reordered because the group changed.");
    }
    const byId = new Map(favorites.map((favorite) => [favorite.id, favorite]));
    const reordered = ids.map((id) => byId.get(id));
    if (reordered.some((favorite) => !favorite)) {
      throw new Error("Favorites could not be reordered because the group changed.");
    }
    reordered.forEach((favorite, index) => {
      if (favorite) favorite.order = index;
    });
    await this.store.saveData(data);
  }

  async moveFavoriteToGroup(id: string, groupId: string): Promise<void> {
    const data = await this.listData();
    const favorite = data.favorites.find((item) => item.id === id);
    if (!favorite) throw new Error("Favorite no longer exists.");
    const currentGroupId = favorite.groupId ?? getUngroupedGroupId();
    const destination = data.groups.some((group) => group.id === groupId) ? groupId : getUngroupedGroupId();
    if (destination === currentGroupId) return;
    favorite.groupId = destination;
    favorite.order = this.nextFavoriteOrder(data, destination);
    favorite.updatedAt = this.now().toISOString();
    await this.store.saveData(data);
  }

  async move(id: string, offset: -1 | 1): Promise<void> {
    const favorites = await this.list();
    const index = favorites.findIndex((favorite) => favorite.id === id);
    const destination = index + offset;
    if (index < 0 || destination < 0 || destination >= favorites.length) return;
    const data = await this.listData();
    const sameGroup = (favorites[index].groupId ?? getUngroupedGroupId()) === (favorites[destination].groupId ?? getUngroupedGroupId());
    if (!sameGroup) return;
    const groupId = favorites[index].groupId ?? getUngroupedGroupId();
    const groupFavorites = data.favorites.filter((favorite) => (favorite.groupId ?? getUngroupedGroupId()) === groupId);
    const byId = new Map(groupFavorites.map((favorite) => [favorite.id, favorite]));
    const targetIndex = groupFavorites.findIndex((favorite) => favorite.id === id);
    const targetDestination = targetIndex + offset;
    if (targetIndex < 0 || targetDestination < 0 || targetDestination >= groupFavorites.length) return;
    [groupFavorites[targetIndex], groupFavorites[targetDestination]] = [groupFavorites[targetDestination], groupFavorites[targetIndex]];
    groupFavorites.forEach((favorite, newIndex) => {
      const current = data.favorites.find((item) => item.id === favorite.id);
      if (current) current.order = newIndex;
    });
    await this.store.saveData(data);
  }
}
