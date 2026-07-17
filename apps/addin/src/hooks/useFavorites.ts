import { useCallback, useState } from "react";

const FAVORITES_STORAGE_KEY = "slidebrary.favorite-slide-ids";

const readFavorites = (): Set<string> => {
  try {
    const value: unknown = JSON.parse(globalThis.localStorage.getItem(FAVORITES_STORAGE_KEY) ?? "[]");
    if (!Array.isArray(value)) {
      return new Set();
    }

    return new Set(value.filter((item): item is string => typeof item === "string"));
  } catch {
    return new Set();
  }
};

export interface FavoritesState {
  favoriteIds: ReadonlySet<string>;
  toggleFavorite(id: string): void;
}

export const useFavorites = (): FavoritesState => {
  const [favoriteIds, setFavoriteIds] = useState<Set<string>>(readFavorites);

  const toggleFavorite = useCallback((id: string): void => {
    setFavoriteIds((current) => {
      const next = new Set(current);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }

      globalThis.localStorage.setItem(FAVORITES_STORAGE_KEY, JSON.stringify([...next]));
      return next;
    });
  }, []);

  return { favoriteIds, toggleFavorite };
};
