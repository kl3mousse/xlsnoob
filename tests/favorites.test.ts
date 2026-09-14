/// <reference types="node" />

import assert from "node:assert/strict";
import test from "node:test";
import { Favorite, getUngroupedGroupId } from "../src/favorites/favorite";
import { deriveFavoriteLocation, deriveFavoriteName, filterFavorites, normalizeFavoriteUrl } from "../src/favorites/favorite-utils";
import { FavoritesService } from "../src/favorites/favorites-service";
import { LocalStorageFavoritesStore } from "../src/favorites/local-storage-favorites-store";
import { getContainingFolderUrl, getExcelOpenUrl } from "../src/office/files";

const v1StorageKey = "xlsnoob.favorites.v1";
const v2StorageKey = "xlsnoob.favorites.v2";

class MemoryStorage {
  private readonly values = new Map<string, string>();

  getItem(key: string): string | null {
    return this.values.get(key) ?? null;
  }

  setItem(key: string, value: string): void {
    this.values.set(key, value);
  }
}

function createFavorite(overrides: Partial<Favorite> = {}): Favorite {
  return {
    id: "favorite-1",
    name: "Finance Forecast",
    url: "https://example.com/docs/forecast.xlsx",
    addedAt: "2026-01-01T00:00:00.000Z",
    order: 0,
    ...overrides,
  };
}

test("empty and malformed storage produce an empty list", async () => {
  const storage = new MemoryStorage();
  const store = new LocalStorageFavoritesStore(storage);
  assert.deepEqual(await store.list(), []);

  storage.setItem(v1StorageKey, "{not json");
  assert.deepEqual(await store.list(), []);
});

test("v1 favorites migrate to v2 grouped data with ungrouped defaults", async () => {
  const storage = new MemoryStorage();
  const legacy = [
    createFavorite({ id: "f-1", name: "Legacy one" }),
    createFavorite({ id: "f-2", name: "Legacy two", groupId: "custom-group" }),
  ];
  storage.setItem(v1StorageKey, JSON.stringify(legacy));

  const store = new LocalStorageFavoritesStore(storage);
  const data = await store.listData();

  assert.equal(data.version, 2);
  assert.deepEqual(data.groups.map((group: { name: string }) => group.name), ["Ungrouped"]);
  assert.deepEqual(data.favorites.map(({ id, groupId, order }: { id: string; groupId?: string; order?: number }) => ({ id, groupId, order })), [
    { id: "f-1", groupId: getUngroupedGroupId(), order: 0 },
    { id: "f-2", groupId: getUngroupedGroupId(), order: 1 },
  ]);
  assert.equal(storage.getItem(v2StorageKey)?.includes('"version":2'), true);
});

test("malformed items are skipped without discarding valid favorites", async () => {
  const storage = new MemoryStorage();
  const valid = createFavorite({ futureGroup: "leadership" });
  storage.setItem(v2StorageKey, JSON.stringify({ version: 2, groups: [{ id: getUngroupedGroupId(), name: "Ungrouped", order: 0, collapsed: false }], favorites: [valid, { id: 4 }, null] }));
  const favorites = await new LocalStorageFavoritesStore(storage).list();
  assert.equal(favorites.length, 1);
  assert.deepEqual(favorites[0].id, valid.id);
  assert.equal(favorites[0].futureGroup, "leadership");
  assert.equal(favorites[0].groupId, getUngroupedGroupId());
});

test("v2 data restores the Ungrouped destination before adding a favorite", async () => {
  const storage = new MemoryStorage();
  storage.setItem(v2StorageKey, JSON.stringify({
    version: 2,
    groups: [{ id: "finance", name: "Finance", order: 0, collapsed: false }],
    favorites: [],
  }));
  const service = new FavoritesService(
    new LocalStorageFavoritesStore(storage),
    () => "new-favorite",
    () => new Date("2026-02-03T04:05:06.000Z"),
  );

  const favorite = await service.add("https://example.com/docs/new-model.xlsx");
  const data = await service.listData();

  assert.equal(favorite.groupId, getUngroupedGroupId());
  assert.equal(data.groups.some(({ id }) => id === getUngroupedGroupId()), true);
  assert.equal(data.favorites[0].groupId, getUngroupedGroupId());
});

test("service adds, deduplicates, edits, reorders, moves groups, and removes favorites", async () => {
  const storage = new MemoryStorage();
  const service = new FavoritesService(
    new LocalStorageFavoritesStore(storage),
    (() => {
      let id = 0;
      return () => `favorite-${++id}`;
    })(),
    () => new Date("2026-02-03T04:05:06.000Z"),
  );

  const first = await service.add("https://EXAMPLE.com/docs/forecast.xlsx#sheet=1", "Forecast", "Weekly");
  const duplicate = await service.add("https://example.com/docs/forecast.xlsx", "Duplicate");
  assert.equal(duplicate.id, first.id);
  assert.equal((await service.list()).length, 1);

  const second = await service.add("https://example.com/docs/headcount.xlsx");
  const third = await service.add("https://example.com/docs/template.xlsx");
  void third;
  const group = await service.addGroup("Finance");
  await service.moveFavoriteToGroup(first.id, group.id);
  await service.moveFavoriteToGroup(second.id, group.id);
  await service.reorderGroupFavorites(group.id, [second.id, first.id]);
  assert.deepEqual((await service.list()).filter(({ groupId }) => groupId === group.id).map(({ id }) => id), [second.id, first.id]);

  await service.update(first.id, {
    name: "Finance Forecast v2",
    comment: "Reference",
    url: first.url,
  });
  const groupFavoritesAfterUpdate = (await service.list()).filter(({ groupId }) => groupId === group.id);
  assert.equal(groupFavoritesAfterUpdate[0].id, second.id);
  assert.equal(groupFavoritesAfterUpdate[1].id, first.id);
  assert.equal(groupFavoritesAfterUpdate[1].name, "Finance Forecast v2");
  assert.equal(groupFavoritesAfterUpdate[1].comment, "Reference");
  assert.equal(groupFavoritesAfterUpdate[1].updatedAt, "2026-02-03T04:05:06.000Z");

  await service.reorderGroupFavorites(group.id, [first.id, second.id]);
  assert.deepEqual((await service.list()).filter(({ groupId }) => groupId === group.id).map(({ id }) => id), [first.id, second.id]);
  await service.moveGroup(group.id, 1);
  assert.deepEqual((await service.listGroups()).map(({ id }) => id), [getUngroupedGroupId(), group.id]);

  await service.remove(first.id);
  assert.deepEqual((await service.list()).filter(({ groupId }) => groupId === group.id).map(({ id }) => id), [second.id]);
});

test("deleting a custom group moves its favorites back to Ungrouped in order", async () => {
  const storage = new MemoryStorage();
  const service = new FavoritesService(
    new LocalStorageFavoritesStore(storage),
    (() => {
      let id = 0;
      return () => `favorite-${++id}`;
    })(),
    () => new Date("2026-02-03T04:05:06.000Z"),
  );

  const ungrouped = await service.add("https://example.com/docs/summary.xlsx");
  const first = await service.add("https://example.com/docs/forecast.xlsx");
  const second = await service.add("https://example.com/docs/headcount.xlsx");
  const group = await service.addGroup("Finance");
  await service.moveFavoriteToGroup(first.id, group.id);
  await service.moveFavoriteToGroup(second.id, group.id);

  await service.deleteGroup(group.id);

  assert.deepEqual((await service.listGroups()).map(({ id }) => id), [getUngroupedGroupId()]);
  assert.deepEqual(
    (await service.list()).map(({ id, groupId }) => ({ id, groupId })),
    [
      { id: ungrouped.id, groupId: getUngroupedGroupId() },
      { id: first.id, groupId: getUngroupedGroupId() },
      { id: second.id, groupId: getUngroupedGroupId() },
    ],
  );
});

test("URL helpers normalize conservatively and derive decoded names", () => {
  assert.equal(
    normalizeFavoriteUrl(" HTTPS://Example.COM:443/docs/Quarterly%20Review.xlsx/#sheet=Summary "),
    "https://example.com/docs/Quarterly%20Review.xlsx",
  );
  assert.equal(
    deriveFavoriteName("https://example.com/docs/Quarterly%20Review.xlsx?download=1"),
    "Quarterly Review",
  );
  assert.equal(
    deriveFavoriteName("https://example.com/docs/S4_Hosting_EA_Update_v1.0.xlsm"),
    "S4 Hosting EA Update v1.0",
  );
  assert.equal(normalizeFavoriteUrl("not a URL "), "not a URL");
});

test("location helper produces short conservative hints", () => {
  assert.equal(deriveFavoriteLocation("https://danone-my.sharepoint.com/personal/user/Documents/forecast.xlsx"), "Personal OneDrive");
  assert.equal(deriveFavoriteLocation("https://danone.sharepoint.com/sites/WW-Cybersecurity/Shared/forecast.xlsx"), "WW-Cybersecurity");
  assert.equal(deriveFavoriteLocation("https://files.example.com/workbooks/file.xlsx"), "files.example.com");
});

test("Excel open URL launches the desktop client in edit mode", () => {
  assert.equal(
    getExcelOpenUrl("https://example.sharepoint.com/sites/team/Shared%20Documents/workbook.xlsx"),
    "ms-excel:ofe|u|https://example.sharepoint.com/sites/team/Shared%20Documents/workbook.xlsx",
  );
  assert.throws(() => getExcelOpenUrl("file:///Users/example/workbook.xlsx"), /HTTP or HTTPS/);
});

test("containing folder URL drops filename, query, and hash", () => {
  assert.equal(
    getContainingFolderUrl("https://example.sharepoint.com/sites/team/Shared%20Documents/workbook.xlsx?web=1#sheet=Summary"),
    "https://example.sharepoint.com/sites/team/Shared%20Documents/",
  );
});

test("filter searches name, comment, and URL without changing manual order", () => {
  const favorites = [
    createFavorite({ id: "1", name: "Finance Forecast", comment: "Weekly" }),
    createFavorite({ id: "2", name: "Architecture Principles", comment: "Reference", url: "https://example.com/architecture.xlsx" }),
    createFavorite({ id: "3", name: "Hiring Tracker", url: "https://example.com/templates/hiring-tracker.xlsx" }),
  ];
  assert.deepEqual(filterFavorites(favorites, "reference").map(({ id }) => id), ["2"]);
  assert.deepEqual(filterFavorites(favorites, "FORECAST").map(({ id }) => id), ["1"]);
  assert.deepEqual(filterFavorites(favorites, "templates").map(({ id }) => id), ["3"]);
  assert.deepEqual(filterFavorites(favorites, "").map(({ id }) => id), ["1", "2", "3"]);
});
