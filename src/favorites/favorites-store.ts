import { Favorite, FavoritesData } from "./favorite";

export interface FavoritesStore {
  list(): Promise<Favorite[]>;
  listData(): Promise<FavoritesData>;
  saveData(data: FavoritesData): Promise<void>;
  add(favorite: Favorite): Promise<void>;
  update(favorite: Favorite): Promise<void>;
  remove(id: string): Promise<void>;
  reorder(ids: string[]): Promise<void>;
}
