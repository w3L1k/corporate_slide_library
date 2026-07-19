import { useCallback, useState } from "react";

const RECENT_SLIDES_STORAGE_KEY = "slidebrary.recent-slide-ids";
const MAX_RECENT_SLIDES = 12;

const normalizeIds = (value: unknown): string[] => {
  if (!Array.isArray(value)) {
    return [];
  }

  const uniqueIds = new Set<string>();
  for (const item of value) {
    if (typeof item === "string" && item.trim()) {
      uniqueIds.add(item);
    }
  }

  return [...uniqueIds].slice(0, MAX_RECENT_SLIDES);
};

const readRecentIds = (): string[] => {
  try {
    return normalizeIds(
      JSON.parse(globalThis.localStorage.getItem(RECENT_SLIDES_STORAGE_KEY) ?? "[]")
    );
  } catch {
    return [];
  }
};

const persistRecentIds = (ids: readonly string[]): void => {
  try {
    globalThis.localStorage.setItem(RECENT_SLIDES_STORAGE_KEY, JSON.stringify(ids));
  } catch {
    // A blocked storage API should not make insertion or browsing unusable.
  }
};

export interface RecentSlidesState {
  recentIds: readonly string[];
  recordRecent(ids: readonly string[]): void;
}

export const useRecentSlides = (): RecentSlidesState => {
  const [recentIds, setRecentIds] = useState<string[]>(readRecentIds);

  const recordRecent = useCallback((ids: readonly string[]): void => {
    const insertedIds = normalizeIds(ids);
    if (insertedIds.length === 0) {
      return;
    }

    setRecentIds((current) => {
      const insertedIdSet = new Set(insertedIds);
      const next = [
        ...insertedIds,
        ...current.filter((id) => !insertedIdSet.has(id))
      ].slice(0, MAX_RECENT_SLIDES);
      persistRecentIds(next);
      return next;
    });
  }, []);

  return { recentIds, recordRecent };
};
