import { getCurrentWorkbookUrl, openContainingFolderUrl, openWorkbookUrl, requireWebUrl } from "../office/files";
import { setStatus, showStatusToast } from "../ui/status";
import { Favorite, FavoriteGroup, getUngroupedGroupId } from "./favorite";
import { deriveFavoriteLocation, deriveFavoriteName, filterFavorites, normalizeFavoriteUrl } from "./favorite-utils";
import { FavoritesService } from "./favorites-service";
import { LocalStorageFavoritesStore } from "./local-storage-favorites-store";

const service = new FavoritesService(new LocalStorageFavoritesStore());
let favorites: Favorite[] = [];
let groups: FavoriteGroup[] = [];
let currentWorkbookUrl = "";
let editingId: string | undefined;
let pendingFavorite: Favorite | undefined;
let removingId: string | undefined;
let searchQuery = "";
let fieldIdSequence = 0;

function element<T extends HTMLElement>(id: string): T {
  return document.getElementById(id) as T;
}

function reportError(error: unknown): void {
  setStatus(error instanceof Error ? error.message : String(error));
}

function showDialogInput(title: string, defaultValue = "", placeholder = "", options?: Array<{ label: string; value: string }>): Promise<string | null> {
  return new Promise((resolve) => {
    const backdrop = document.createElement("div");
    backdrop.className = "favorite-dialog-backdrop";
    const dialog = document.createElement("div");
    dialog.className = "favorite-dialog";
    const heading = document.createElement("h3");
    heading.textContent = title;
    const field = options && options.length > 0
      ? document.createElement("select")
      : document.createElement("input");
    field.className = "control control-compact";
    if (field instanceof HTMLInputElement) {
      field.type = "text";
      field.value = defaultValue;
      field.placeholder = placeholder;
    } else {
      options!.forEach(({ label, value }) => {
        const option = document.createElement("option");
        option.textContent = label;
        option.value = value;
        if (value === defaultValue) option.selected = true;
        field.append(option);
      });
    }
    const actions = document.createElement("div");
    actions.className = "favorite-dialog-actions";
    const cancel = createButton("Cancel", () => {
      backdrop.remove();
      resolve(null);
    }, "btn btn-ghost btn-sm");
    const submit = createButton("OK", () => {
      const value = field instanceof HTMLInputElement ? field.value.trim() : field.value;
      backdrop.remove();
      resolve(value || null);
    }, "btn btn-primary btn-sm");
    actions.append(cancel, submit);
    dialog.append(heading, field, actions);
    backdrop.append(dialog);
    backdrop.addEventListener("click", (event) => {
      if (event.target === backdrop) {
        backdrop.remove();
        resolve(null);
      }
    });
    field.addEventListener("keydown", (event) => {
      const keyboardEvent = event as KeyboardEvent;
      if (keyboardEvent.key === "Enter") {
        keyboardEvent.preventDefault();
        submit.click();
      }
      if (keyboardEvent.key === "Escape") {
        keyboardEvent.preventDefault();
        cancel.click();
      }
    });
    document.body.append(backdrop);
    queueMicrotask(() => (field as HTMLInputElement | HTMLSelectElement).focus());
  });
}

function showDialogConfirm(title: string, message: string, confirmLabel: string): Promise<boolean> {
  return new Promise((resolve) => {
    const backdrop = document.createElement("div");
    backdrop.className = "favorite-dialog-backdrop";
    const dialog = document.createElement("div");
    dialog.className = "favorite-dialog";
    const heading = document.createElement("h3");
    heading.textContent = title;
    const body = document.createElement("p");
    body.textContent = message;
    const actions = document.createElement("div");
    actions.className = "favorite-dialog-actions";
    const cancel = createButton("Cancel", () => {
      backdrop.remove();
      resolve(false);
    }, "btn btn-ghost btn-sm");
    const confirm = createButton(confirmLabel, () => {
      backdrop.remove();
      resolve(true);
    }, "btn btn-primary btn-sm");
    actions.append(cancel, confirm);
    dialog.append(heading, body, actions);
    backdrop.append(dialog);
    backdrop.addEventListener("click", (event) => {
      if (event.target === backdrop) {
        backdrop.remove();
        resolve(false);
      }
    });
    document.body.append(backdrop);
    queueMicrotask(() => confirm.focus());
  });
}

function run(action: () => Promise<void>): void {
  action().catch(reportError);
}

function createButton(label: string, action: () => void, className = "btn btn-ghost btn-sm"): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = className;
  button.textContent = label;
  button.addEventListener("click", action);
  return button;
}

function getGroupName(groupId: string | undefined): string {
  return groups.find((group) => group.id === (groupId ?? getUngroupedGroupId()))?.name ?? "Ungrouped";
}

function updateCurrentState(): void {
  const current = currentWorkbookUrl
    ? favorites.find((favorite) => normalizeFavoriteUrl(favorite.url) === normalizeFavoriteUrl(currentWorkbookUrl))
    : undefined;
  const button = element<HTMLButtonElement>("addCurrentFavorite");
  button.textContent = current ? "★ Saved" : "☆ Add";
  button.title = current ? "Current workbook saved in Favorites" : "Add current workbook to Favorites";
  button.setAttribute("aria-label", current ? "Current workbook saved in Favorites" : "Add current workbook to Favorites");
  button.classList.toggle("is-favorited", Boolean(current));
  button.dataset.favoriteId = current?.id || "";
  const count = element<HTMLElement>("favoritesCount");
  count.textContent = String(favorites.length);
}

function createField(labelText: string, control: HTMLElement): HTMLLabelElement {
  const label = document.createElement("label");
  const text = document.createElement("span");
  label.className = "favorite-field";
  text.textContent = labelText;
  if (control instanceof HTMLInputElement || control instanceof HTMLSelectElement || control instanceof HTMLTextAreaElement) {
    const fieldId = `favorite-field-${++fieldIdSequence}`;
    control.id = fieldId;
    label.htmlFor = fieldId;
  }
  label.append(text, control);
  return label;
}

function createEditForm(favorite: Favorite): HTMLElement {
  const isNew = !favorite.id;
  const form = document.createElement("form");
  form.className = "favorite-edit";
  const title = document.createElement("h3");
  title.textContent = isNew ? "Add to favorites" : "Edit favorite";
  const name = document.createElement("input");
  name.className = "control control-compact";
  name.value = favorite.name;
  name.required = true;
  const note = document.createElement("input");
  note.className = "control control-compact";
  note.value = favorite.comment || "";
  note.placeholder = "Optional note";
  note.maxLength = 120;
  const url = document.createElement("input");
  url.className = "control control-compact";
  url.type = "url";
  url.value = favorite.url;
  url.required = true;
  const actions = document.createElement("div");
  actions.className = "favorite-form-actions";
  const cancel = createButton("Cancel", () => {
    editingId = undefined;
    pendingFavorite = undefined;
    removingId = undefined;
    renderFavorites();
  });
  const submit = createButton(isNew ? "Add" : "Save", () => undefined, "btn btn-primary btn-sm");
  submit.type = "submit";
  actions.append(cancel, submit);
  form.append(title, createField("Name", name), createField("Note", note));
  if (isNew) {
    const location = document.createElement("div");
    location.className = "favorite-form-location";
    location.textContent = deriveFavoriteLocation(favorite.url);
    location.title = favorite.url;
    form.append(createField("Location", location));
  } else {
    form.append(createField("Location", url));
  }
  form.append(actions);
  if (!isNew) {
    const removeArea = document.createElement("div");
    removeArea.className = "favorite-remove-area";
    if (removingId === favorite.id) {
      const prompt = document.createElement("span");
      prompt.textContent = "Remove this favorite?";
      removeArea.append(
        prompt,
        createButton("Cancel", () => {
          removingId = undefined;
          renderFavorites();
        }),
        createButton("Remove", () => run(async () => {
          await service.remove(favorite.id);
          editingId = undefined;
          removingId = undefined;
          await refreshFavorites();
          showStatusToast("✓ Removed from favorites");
        }), "btn btn-danger btn-sm"),
      );
    } else {
      removeArea.append(createButton("Remove favorite", () => {
        removingId = favorite.id;
        renderFavorites();
      }, "btn btn-link btn-danger-text btn-sm"));
    }
    form.append(removeArea);
  }
  form.addEventListener("submit", (event) => {
    event.preventDefault();
    run(async () => {
      requireWebUrl(url.value);
      if (favorite.id) {
        await service.update(favorite.id, { name: name.value, comment: note.value, url: url.value });
      } else {
        await service.add(url.value, name.value, note.value);
      }
      editingId = undefined;
      pendingFavorite = undefined;
      await refreshFavorites();
      showStatusToast(favorite.id ? "✓ Favorite updated" : "✓ Added to favorites");
    });
  });
  queueMicrotask(() => name.focus());
  return form;
}

async function promptMoveFavoriteToGroup(id: string): Promise<string | null> {
  const currentFavorite = favorites.find((favorite) => favorite.id === id);
  const currentGroup = currentFavorite?.groupId ?? getUngroupedGroupId();
  const options = [
    { label: "Ungrouped", value: getUngroupedGroupId() },
    ...groups
      .slice()
      .sort((left, right) => left.order - right.order)
      .filter((group) => group.id !== getUngroupedGroupId())
      .map((group) => ({ label: group.name, value: group.id })),
  ];
  const target = await showDialogInput("Move to group", currentGroup, "", options);
  if (!target) return null;
  return options.find((option) => option.value === target)?.value ?? null;
}

function promptCreateGroup(): void {
  run(async () => {
    const value = await showDialogInput("New group", "", "Group name");
    if (!value) return;
    const name = value.trim();
    if (!name) return;
    await service.addGroup(name);
    await refreshFavorites();
    showStatusToast("✓ Group created");
  });
}

function promptRenameGroup(groupId: string): void {
  const group = groups.find((item) => item.id === groupId);
  if (!group) return;
  run(async () => {
    const value = await showDialogInput("Rename group", group.name, "Group name");
    if (!value) return;
    const name = value.trim();
    if (!name) return;
    await service.renameGroup(groupId, name);
    await refreshFavorites();
  });
}

function createFavoriteMenu(favorite: Favorite, indexInGroup: number, groupId: string): HTMLDetailsElement {
  const menu = document.createElement("details");
  menu.className = "action-menu";
  const menuTrigger = document.createElement("summary");
  menuTrigger.className = "icon-action";
  menuTrigger.textContent = "⋯";
  menuTrigger.title = `More actions for ${favorite.name}`;
  menuTrigger.setAttribute("role", "button");
  menuTrigger.setAttribute("aria-label", `More actions for ${favorite.name}`);
  menuTrigger.setAttribute("aria-haspopup", "menu");
  menuTrigger.setAttribute("aria-expanded", "false");
  const menuItems = document.createElement("div");
  menuItems.className = "action-menu-items";
  menuItems.setAttribute("role", "menu");
  const createMenuItem = (label: string, action: () => void, disabled = false): HTMLButtonElement => {
    const item = createButton(label, () => {
      menu.open = false;
      action();
    }, "action-menu-item");
    item.disabled = disabled;
    item.setAttribute("role", "menuitem");
    return item;
  };
  const groupFavorites = favorites.filter((item) => (item.groupId ?? getUngroupedGroupId()) === groupId);
  const isFirst = indexInGroup === 0;
  const isLast = indexInGroup === groupFavorites.length - 1;
  menuItems.append(
    createMenuItem("Open containing folder", () => {
      try {
        openContainingFolderUrl(favorite.url);
        showStatusToast("✓ Opened containing folder");
      } catch (error) {
        reportError(error);
      }
    }),
    createMenuItem("Edit", () => {
      editingId = favorite.id;
      removingId = undefined;
      renderFavorites();
    }),
    createMenuItem("Move to group >", () => {
      run(async () => {
        const destination = await promptMoveFavoriteToGroup(favorite.id);
        if (!destination) return;
        await service.moveFavoriteToGroup(favorite.id, destination);
        await refreshFavorites();
        showStatusToast(`✓ Moved to ${getGroupName(destination)}`);
      });
    }),
    createMenuItem("Move up", () => run(async () => {
      await service.move(favorite.id, -1);
      await refreshFavorites();
    }), isFirst),
    createMenuItem("Move down", () => run(async () => {
      await service.move(favorite.id, 1);
      await refreshFavorites();
    }), isLast),
    createMenuItem("Remove", () => {
      editingId = favorite.id;
      removingId = favorite.id;
      renderFavorites();
    }),
  );
  menu.append(menuTrigger, menuItems);
  menu.addEventListener("toggle", () => {
    menuTrigger.setAttribute("aria-expanded", String(menu.open));
    if (!menu.open) return;
    document.querySelectorAll<HTMLDetailsElement>(".action-menu[open]").forEach((other) => {
      if (other !== menu) other.open = false;
    });
  });
  return menu;
}

function createFavoriteCard(favorite: Favorite, groupId: string, indexInGroup: number): HTMLElement {
  const card = document.createElement("article");
  card.className = "favorite-card";
  card.dataset.favoriteId = favorite.id;
  if (editingId === favorite.id) {
    card.append(createEditForm(favorite));
    return card;
  }

  const main = document.createElement("button");
  main.type = "button";
  main.className = "favorite-main";
  main.title = `Open ${favorite.name}\n${favorite.url}`;
  main.setAttribute("aria-label", `Open ${favorite.name}`);
  const content = document.createElement("div");
  content.className = "favorite-main-content";
  if (favorite.comment) {
    const note = document.createElement("span");
    note.className = "favorite-note";
    note.textContent = favorite.comment;
    content.append(note);
  }
  const titleRow = document.createElement("div");
  titleRow.className = "favorite-title-row";
  const name = document.createElement("strong");
  name.textContent = favorite.name;
  titleRow.append(name);
  const location = document.createElement("span");
  location.className = "favorite-location";
  location.textContent = deriveFavoriteLocation(favorite.url);
  location.title = favorite.url;
  content.append(titleRow, location);
  main.append(content);
  main.addEventListener("click", () => {
    try {
      openWorkbookUrl(favorite.url);
      showStatusToast("Opening in Excel…");
    } catch (error) {
      reportError(error);
    }
  });
  const menu = createFavoriteMenu(favorite, indexInGroup, groupId);
  card.append(main, menu);
  return card;
}

function createGroupSection(group: FavoriteGroup, groupFavorites: Favorite[]): HTMLElement {
  const section = document.createElement("section");
  section.className = "favorite-group surface";
  const pendingInGroup = pendingFavorite && group.id === getUngroupedGroupId() ? pendingFavorite : undefined;
  const header = document.createElement("button");
  header.type = "button";
  header.className = "favorite-group-header";
  header.setAttribute("aria-expanded", String(!group.collapsed));
  const chevron = document.createElement("span");
  chevron.className = "favorite-group-chevron";
  chevron.textContent = group.collapsed ? "▸" : "▾";
  const label = document.createElement("span");
  label.className = "favorite-group-label";
  label.textContent = group.name;
  const count = document.createElement("span");
  count.className = "favorite-group-count";
  count.textContent = `· ${groupFavorites.length}`;
  header.append(chevron, label, count);
  header.addEventListener("click", () => {
    run(async () => {
      await service.updateGroupCollapsed(group.id, !group.collapsed);
      await refreshFavorites();
    });
  });
  const menu = document.createElement("details");
  menu.className = "action-menu favorite-group-menu";
  const menuTrigger = document.createElement("summary");
  menuTrigger.className = "icon-action";
  menuTrigger.textContent = "⋯";
  menuTrigger.setAttribute("aria-label", `Group actions for ${group.name}`);
  const menuItems = document.createElement("div");
  menuItems.className = "action-menu-items";
  const addItem = (label: string, action: () => void): HTMLButtonElement => {
    const button = createButton(label, () => {
      menu.open = false;
      action();
    }, "action-menu-item");
    button.setAttribute("role", "menuitem");
    return button;
  };
  menuItems.append(
    addItem("Rename", () => promptRenameGroup(group.id)),
    addItem("Move up", () => run(async () => {
      await service.moveGroup(group.id, -1);
      await refreshFavorites();
    })),
    addItem("Move down", () => run(async () => {
      await service.moveGroup(group.id, 1);
      await refreshFavorites();
    })),
  );
  if (group.id !== getUngroupedGroupId()) {
    menuItems.append(addItem("Delete group", () => run(async () => {
      const confirmed = await showDialogConfirm("Delete group", `Move favorites from ${group.name} to Ungrouped and delete this group?`, "Delete");
      if (!confirmed) return;
      await service.deleteGroup(group.id);
      await refreshFavorites();
      showStatusToast("✓ Group deleted");
    })));
  }
  menu.append(menuTrigger, menuItems);
  menu.addEventListener("toggle", () => {
    menuTrigger.setAttribute("aria-expanded", String(menu.open));
    if (!menu.open) return;
    document.querySelectorAll<HTMLDetailsElement>(".action-menu[open]").forEach((other) => {
      if (other !== menu) other.open = false;
    });
  });
  const list = document.createElement("div");
  list.className = "favorite-group-items";
  if (!pendingInGroup && !searchQuery && group.collapsed) {
    list.hidden = true;
  }
  if (pendingInGroup) {
    list.append(createFavoriteCard(pendingInGroup, group.id, 0));
  }
  if (!group.collapsed || searchQuery || pendingInGroup) {
    const items = groupFavorites.map((favorite, index) => createFavoriteCard(favorite, group.id, index));
    list.append(...items);
  }
  const headerRow = document.createElement("div");
  headerRow.className = "favorite-group-header-row";
  headerRow.append(header, menu);
  section.append(headerRow, list);
  return section;
}

function renderFavorites(): void {
  const query = searchQuery.trim();
  const filteredFavorites = filterFavorites(favorites, query, groups);
  const favoritesToGroup = query ? filteredFavorites : favorites;
  const itemsByGroup = new Map<string, Favorite[]>();
  for (const favorite of favoritesToGroup) {
    const key = favorite.groupId ?? getUngroupedGroupId();
    const groupItems = itemsByGroup.get(key) ?? [];
    groupItems.push(favorite);
    itemsByGroup.set(key, groupItems);
  }
  const visibleGroups = groups
    .slice()
    .sort((left, right) => left.order - right.order)
    .filter((group) => {
      if (query) {
        const containsPendingFavorite = Boolean(pendingFavorite) && group.id === getUngroupedGroupId();
        return containsPendingFavorite || (itemsByGroup.get(group.id)?.length ?? 0) > 0;
      }
      return true;
    });
  const sections = visibleGroups.map((group) => createGroupSection(group, itemsByGroup.get(group.id) ?? []));
  element("favoritesList").replaceChildren(...sections);
  const empty = element("emptyFavorites");
  empty.hidden = sections.length > 0 || Boolean(pendingFavorite);
  empty.textContent = favorites.length ? "No favorites match your search." : "No favorites yet.";
  updateCurrentState();
}

async function refreshFavorites(): Promise<void> {
  const data = await service.listData();
  favorites = data.favorites;
  groups = data.groups;
  renderFavorites();
}

async function addCurrentWorkbook(): Promise<void> {
  if (!currentWorkbookUrl) {
    throw new Error("Save this workbook to SharePoint, OneDrive, or another supported web location before adding it.");
  }
  requireWebUrl(currentWorkbookUrl);
  const existing = await service.findByUrl(currentWorkbookUrl);
  if (existing) {
    editingId = existing.id;
    pendingFavorite = undefined;
    removingId = undefined;
    renderFavorites();
    queueMicrotask(() => {
      const card = document.querySelector<HTMLElement>(`[data-favorite-id="${existing.id}"]`);
      card?.scrollIntoView({ behavior: "smooth", block: "nearest" });
    });
    return;
  }
  pendingFavorite = {
    id: "",
    name: deriveFavoriteName(currentWorkbookUrl),
    url: currentWorkbookUrl,
    addedAt: "",
    groupId: getUngroupedGroupId(),
    order: 0,
  };
  editingId = "";
  renderFavorites();
}

async function loadCurrentWorkbook(): Promise<void> {
  try {
    currentWorkbookUrl = await getCurrentWorkbookUrl();
  } catch (error) {
    currentWorkbookUrl = "";
    reportError(error);
  }
  updateCurrentState();
}

export function initializeFavorites(): void {
  const addCurrent = element<HTMLButtonElement>("addCurrentFavorite");
  addCurrent.addEventListener("click", () => run(addCurrentWorkbook));
  const addGroup = element<HTMLButtonElement>("addFavoriteGroup");
  addGroup.addEventListener("click", promptCreateGroup);
  const search = element<HTMLInputElement>("favoriteSearch");
  search.addEventListener("input", (event) => {
    searchQuery = (event.target as HTMLInputElement).value;
    renderFavorites();
  });
  search.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    const first = filterFavorites(favorites, searchQuery, groups)[0];
    if (!first) return;
    event.preventDefault();
    try {
      openWorkbookUrl(first.url);
      showStatusToast("Opening in Excel…");
    } catch (error) {
      reportError(error);
    }
  });
  document.addEventListener("click", (event) => {
    const target = event.target;
    if (target instanceof Element && target.closest(".action-menu")) return;
    document.querySelectorAll<HTMLDetailsElement>(".action-menu[open]").forEach((menu) => {
      menu.open = false;
    });
  });
  run(async () => {
    await Promise.all([refreshFavorites(), loadCurrentWorkbook()]);
  });
}
