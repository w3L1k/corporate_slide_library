import { useCallback, useEffect, useMemo, useState } from "react";
import type {
  PersonalAssetKind,
  SlideLibraryItem,
  SlideListResponse,
  SlideStatus
} from "@slide-library/shared";
import { PersonalLibrary } from "./components/PersonalLibrary";
import { SkeletonCatalog } from "./components/SkeletonCatalog";
import { SlideCard } from "./components/SlideCard";
import { SlideDetailsDialog } from "./components/SlideDetailsDialog";
import { Toast, type ToastMessage } from "./components/Toast";
import { useDebouncedValue } from "./hooks/useDebouncedValue";
import { useFavorites } from "./hooks/useFavorites";
import {
  slideLibraryApi,
  type SlideLibraryApi,
  type SlideListFilters
} from "./services/api";
import {
  BrowserPowerPointService,
  PowerPointUnavailableError,
  type PowerPointService
} from "./services/powerpoint";
import "./styles.css";

type StatusFilter = SlideStatus | "";
type LibrarySection = "favorites" | "presentations";
type LibraryScope = "public" | "personal";
type ViewMode = "grid" | "list";
type SortOrder = "updated-desc" | "title-asc";

const DEFAULT_STATUS: StatusFilter = "approved";
const DEFAULT_DEBOUNCE_MS = 300;
const MULTI_INSERT_ID = "__multiple__";

interface AppProps {
  api?: SlideLibraryApi;
  powerPointService?: PowerPointService;
  searchDebounceMs?: number;
}

interface CatalogState {
  response: SlideListResponse | null;
  loading: boolean;
  error: string | null;
}

const initialCatalogState: CatalogState = {
  response: null,
  loading: true,
  error: null
};

const getErrorMessage = (error: unknown): string => {
  if (error instanceof Error && error.message.trim()) {
    return error.message;
  }

  return "The slide library could not be loaded.";
};

export function App({
  api = slideLibraryApi,
  powerPointService = new BrowserPowerPointService(),
  searchDebounceMs = DEFAULT_DEBOUNCE_MS
}: AppProps) {
  const [query, setQuery] = useState("");
  const [activeSection, setActiveSection] = useState<LibrarySection>("presentations");
  const [libraryScope, setLibraryScope] = useState<LibraryScope>("public");
  const [personalKind, setPersonalKind] = useState<PersonalAssetKind>("presentation");
  const [viewMode, setViewMode] = useState<ViewMode>("grid");
  const [sortOrder, setSortOrder] = useState<SortOrder>("updated-desc");
  const debouncedQuery = useDebouncedValue(query.trim(), searchDebounceMs);
  const [category, setCategory] = useState("");
  const [status, setStatus] = useState<StatusFilter>(DEFAULT_STATUS);
  const [refreshKey, setRefreshKey] = useState(0);
  const [catalog, setCatalog] = useState<CatalogState>(initialCatalogState);
  const [categories, setCategories] = useState<string[]>([]);
  const [selectedSlide, setSelectedSlide] = useState<SlideLibraryItem | null>(null);
  const [selectedSlideIds, setSelectedSlideIds] = useState<Set<string>>(new Set());
  const [insertingId, setInsertingId] = useState<string | null>(null);
  const [toast, setToast] = useState<ToastMessage | null>(null);
  const { favoriteIds, toggleFavorite } = useFavorites();
  const powerPointUnavailableReason = powerPointService.getUnavailableReason();

  useEffect(() => {
    const controller = new AbortController();
    const filters: SlideListFilters = {};

    if (debouncedQuery) {
      filters.query = debouncedQuery;
    }
    if (category) {
      filters.category = category;
    }
    if (status) {
      filters.status = status;
    }

    queueMicrotask(() => {
      if (!controller.signal.aborted) {
        setCatalog({ response: null, loading: true, error: null });
      }
    });

    void Promise.resolve()
      .then(() => api.listSlides(filters, controller.signal))
      .then((response) => {
        if (controller.signal.aborted) {
          return;
        }

        setCatalog({ response, loading: false, error: null });
        setCategories(response.availableCategories);
      })
      .catch((error: unknown) => {
        if (controller.signal.aborted) {
          return;
        }

        setCatalog({ response: null, loading: false, error: getErrorMessage(error) });
      });

    return () => controller.abort();
  }, [api, category, debouncedQuery, refreshKey, status]);

  const hasActiveFilters = Boolean(query.trim() || category || status !== DEFAULT_STATUS);
  const isEmptyCatalog =
    !hasActiveFilters &&
    catalog.response !== null &&
    catalog.response.items.length === 0 &&
    categories.length === 0;

  const resetFilters = (): void => {
    setQuery("");
    setCategory("");
    setStatus(DEFAULT_STATUS);
  };

  const insertSlide = useCallback(
    async (item: SlideLibraryItem): Promise<void> => {
      if (insertingId !== null) {
        return;
      }

      setInsertingId(item.id);
      try {
        await powerPointService.insertSlide(item.id);
        setToast({ kind: "success", text: "Slide inserted successfully" });
      } catch (error: unknown) {
        if (import.meta.env.DEV && !(error instanceof PowerPointUnavailableError)) {
          console.error("Slide insertion failed", error);
        }

        const message =
          error instanceof PowerPointUnavailableError
            ? error.message
            : `Could not insert slide. ${getErrorMessage(error)}`;
        setToast({ kind: error instanceof PowerPointUnavailableError ? "info" : "error", text: message });
      } finally {
        setInsertingId(null);
      }
    },
    [insertingId, powerPointService]
  );

  const currentSubtitle = useMemo(() => {
    if (status === "approved") {
      return "Approved content, ready to use";
    }
    if (status === "draft") {
      return "Draft content for review";
    }
    if (status === "deprecated") {
      return "Deprecated content — use with care";
    }
    return "All corporate content";
  }, [status]);

  const visibleItems = useMemo(() => {
    const items = [...(catalog.response?.items ?? [])];

    if (sortOrder === "title-asc") {
      return items.sort((left, right) =>
        left.title.localeCompare(right.title, undefined, { sensitivity: "base" })
      );
    }

    return items.sort(
      (left, right) =>
        new Date(right.updatedAt).getTime() - new Date(left.updatedAt).getTime()
    );
  }, [catalog.response?.items, sortOrder]);
  const favoriteItems = useMemo(
    () => visibleItems.filter((item) => favoriteIds.has(item.id)),
    [favoriteIds, visibleItems]
  );
  const selectedItems = useMemo(
    () => visibleItems.filter((item) => selectedSlideIds.has(item.id)),
    [selectedSlideIds, visibleItems]
  );

  const toggleSlideSelection = useCallback((id: string): void => {
    setSelectedSlideIds((current) => {
      const next = new Set(current);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
  }, []);

  const insertSelectedSlides = useCallback(async (): Promise<void> => {
    if (insertingId !== null || selectedItems.length === 0) {
      return;
    }

    setInsertingId(MULTI_INSERT_ID);
    try {
      await powerPointService.insertSlides(selectedItems.map((item) => item.id));
      setSelectedSlideIds(new Set());
      setToast({
        kind: "success",
        text: `${selectedItems.length} ${
          selectedItems.length === 1 ? "slide" : "slides"
        } inserted successfully`
      });
    } catch (error: unknown) {
      if (import.meta.env.DEV && !(error instanceof PowerPointUnavailableError)) {
        console.error("Multiple slide insertion failed", error);
      }

      const message =
        error instanceof PowerPointUnavailableError
          ? error.message
          : `Could not insert selected slides. ${getErrorMessage(error)}`;
      setToast({
        kind: error instanceof PowerPointUnavailableError ? "info" : "error",
        text: message
      });
    } finally {
      setInsertingId(null);
    }
  }, [insertingId, powerPointService, selectedItems]);

  return (
    <div className="app-shell">
      <header className="app-header">
        <div className="brand">
          <button className="profile-button" type="button" aria-label="Открыть профиль">
            <svg viewBox="0 0 24 24">
              <circle cx="12" cy="8" r="3.5" />
              <path d="M5.5 19c.8-3.2 3-5 6.5-5s5.7 1.8 6.5 5" />
            </svg>
          </button>
          <h1>Slidebrary</h1>
        </div>
        <div className="header-actions">
          <span className="header-status">{currentSubtitle}</span>
          <button
            className="icon-button icon-button--refresh"
            type="button"
            onClick={() => setRefreshKey((value) => value + 1)}
            aria-label="Refresh library"
            title="Refresh library"
            disabled={catalog.loading}
          >
            <svg className={catalog.loading ? "is-spinning" : ""} viewBox="0 0 20 20" aria-hidden="true">
              <path d="M15.5 6.2A6 6 0 104 13.8M15.5 6.2V2.8M15.5 6.2h-3.4" />
            </svg>
          </button>
        </div>
      </header>

      <main>
        <nav className="scope-tabs" aria-label="Область библиотеки">
          <button
            className={`scope-tabs__item ${
              libraryScope === "public" ? "scope-tabs__item--active" : ""
            }`}
            type="button"
            onClick={() => setLibraryScope("public")}
            aria-current={libraryScope === "public" ? "page" : undefined}
          >
            Публичное
          </button>
          <button
            className={`scope-tabs__item ${
              libraryScope === "personal" ? "scope-tabs__item--active" : ""
            }`}
            type="button"
            onClick={() => setLibraryScope("personal")}
            aria-current={libraryScope === "personal" ? "page" : undefined}
          >
            Личное
          </button>
        </nav>

        <div className="workspace-layout">
          <aside className="library-sidebar" aria-label="Разделы библиотеки">
            <div className="library-sidebar__rail" aria-hidden="true">
              <span>‹</span>
            </div>
            <div className="library-sidebar__menu">
              <button
                className={`library-sidebar__item ${
                  libraryScope === "public" && activeSection === "favorites"
                    ? "library-sidebar__item--active"
                    : ""
                }`}
                type="button"
                onClick={() => setActiveSection("favorites")}
                disabled={libraryScope === "personal"}
                aria-current={
                  libraryScope === "public" && activeSection === "favorites"
                    ? "page"
                    : undefined
                }
              >
                <span aria-hidden="true">♡</span>
                <strong>Избранное</strong>
              </button>
              <button
                className={`library-sidebar__item ${
                  (libraryScope === "public" && activeSection === "presentations") ||
                  (libraryScope === "personal" && personalKind === "presentation")
                    ? "library-sidebar__item--active"
                    : ""
                }`}
                type="button"
                onClick={() => {
                  if (libraryScope === "personal") {
                    setPersonalKind("presentation");
                  } else {
                    setActiveSection("presentations");
                  }
                }}
                aria-current={
                  (libraryScope === "public" && activeSection === "presentations") ||
                  (libraryScope === "personal" && personalKind === "presentation")
                    ? "page"
                    : undefined
                }
              >
                <span aria-hidden="true">▧</span>
                <strong>Презентации</strong>
              </button>
              <button
                className={`library-sidebar__item ${
                  libraryScope === "personal" && personalKind === "photo"
                    ? "library-sidebar__item--active"
                    : ""
                }`}
                type="button"
                disabled={libraryScope === "public"}
                onClick={() => setPersonalKind("photo")}
                aria-current={
                  libraryScope === "personal" && personalKind === "photo"
                    ? "page"
                    : undefined
                }
              >
                <span aria-hidden="true">▣</span>
                <strong>Фотографии</strong>
              </button>
              <button className="library-sidebar__item" type="button" disabled>
                <span aria-hidden="true">◇</span>
                <strong>Иллюстрации</strong>
              </button>
              <button className="library-sidebar__item" type="button" disabled>
                <span aria-hidden="true">◎</span>
                <strong>Иконки</strong>
              </button>
              <button
                className={`library-sidebar__item ${
                  libraryScope === "personal" && personalKind === "logo"
                    ? "library-sidebar__item--active"
                    : ""
                }`}
                type="button"
                disabled={libraryScope === "public"}
                onClick={() => setPersonalKind("logo")}
                aria-current={
                  libraryScope === "personal" && personalKind === "logo"
                    ? "page"
                    : undefined
                }
              >
                <span aria-hidden="true">A</span>
                <strong>Логотипы</strong>
              </button>
              <button className="library-sidebar__item" type="button" disabled>
                <span aria-hidden="true">⊞</span>
                <strong>Шаблоны</strong>
              </button>
              <button className="library-sidebar__item" type="button" disabled>
                <span aria-hidden="true">✦</span>
                <strong>Продукты</strong>
              </button>
              <button
                className="library-sidebar__item library-sidebar__item--muted"
                type="button"
                disabled
              >
                <span aria-hidden="true">✧</span>
                <strong>ИИ-ассистент</strong>
              </button>
            </div>
          </aside>

          <div className="workspace-content">
        {libraryScope === "personal" ? (
          <PersonalLibrary
            api={api}
            kind={personalKind}
            powerPointService={powerPointService}
            onNotify={setToast}
          />
        ) : null}

        {libraryScope === "public" ? (
          <>
        {powerPointUnavailableReason ? (
          <aside className="environment-notice" aria-label="PowerPoint integration status">
            <span className="environment-notice__icon" aria-hidden="true">
              <svg viewBox="0 0 20 20">
                <circle cx="10" cy="10" r="7.5" />
                <path d="M10 8.5v5M10 5.8v.2" />
              </svg>
            </span>
            <div>
              <strong>Catalog preview mode</strong>
              <p>{powerPointUnavailableReason}</p>
            </div>
          </aside>
        ) : null}

        <section className="catalog-controls" aria-label="Find slides">
          <div className="search-field">
            <svg viewBox="0 0 20 20" aria-hidden="true">
              <circle cx="9" cy="9" r="5.5" />
              <path d="M13 13l4 4" />
            </svg>
            <label className="sr-only" htmlFor="slide-search">
              Search slides
            </label>
            <input
              id="slide-search"
              type="search"
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              placeholder="Search slides…"
              autoComplete="off"
            />
            {query ? (
              <button
                type="button"
                onClick={() => setQuery("")}
                aria-label="Clear search"
                title="Clear search"
              >
                <svg viewBox="0 0 16 16" aria-hidden="true">
                  <path d="M4 4l8 8M12 4l-8 8" />
                </svg>
              </button>
            ) : null}
          </div>

          <div className="filter-row">
            <label>
              <span>Category</span>
              <span className="select-wrap">
                <select value={category} onChange={(event) => setCategory(event.target.value)}>
                  <option value="">All categories</option>
                  {categories.map((availableCategory) => (
                    <option value={availableCategory} key={availableCategory}>
                      {availableCategory}
                    </option>
                  ))}
                </select>
                <svg viewBox="0 0 16 16" aria-hidden="true">
                  <path d="M4 6l4 4 4-4" />
                </svg>
              </span>
            </label>

            <label>
              <span>Status</span>
              <span className="select-wrap">
                <select
                  value={status}
                  onChange={(event) => setStatus(event.target.value as StatusFilter)}
                >
                  <option value="approved">Approved</option>
                  <option value="draft">Draft</option>
                  <option value="deprecated">Deprecated</option>
                  <option value="">All statuses</option>
                </select>
                <svg viewBox="0 0 16 16" aria-hidden="true">
                  <path d="M4 6l4 4 4-4" />
                </svg>
              </span>
            </label>
          </div>

          <div className="results-toolbar">
            <p aria-live="polite" aria-atomic="true">
              {catalog.response ? (
                <>
                  <strong>{catalog.response.total}</strong>{" "}
                  {catalog.response.total === 1 ? "slide" : "slides"}
                </>
              ) : catalog.loading ? (
                "Loading…"
              ) : (
                "—"
              )}
            </p>
            <div className="results-toolbar__actions">
              <label className="sort-control">
                <span className="sr-only">Sort slides</span>
                <select
                  value={sortOrder}
                  onChange={(event) => setSortOrder(event.target.value as SortOrder)}
                  aria-label="Sort slides"
                >
                  <option value="updated-desc">Сначала новые</option>
                  <option value="title-asc">По названию</option>
                </select>
              </label>
              <div className="view-switcher" role="group" aria-label="Режим отображения">
                <button
                  type="button"
                  onClick={() => setViewMode("grid")}
                  aria-label="Плитка"
                  aria-pressed={viewMode === "grid"}
                  title="Плитка"
                >
                  <svg viewBox="0 0 16 16" aria-hidden="true">
                    <rect x="2" y="2" width="4.5" height="4.5" rx="1" />
                    <rect x="9.5" y="2" width="4.5" height="4.5" rx="1" />
                    <rect x="2" y="9.5" width="4.5" height="4.5" rx="1" />
                    <rect x="9.5" y="9.5" width="4.5" height="4.5" rx="1" />
                  </svg>
                </button>
                <button
                  type="button"
                  onClick={() => setViewMode("list")}
                  aria-label="Список"
                  aria-pressed={viewMode === "list"}
                  title="Список"
                >
                  <svg viewBox="0 0 16 16" aria-hidden="true">
                    <path d="M5 3h9M5 8h9M5 13h9" />
                    <circle cx="2" cy="3" r=".7" />
                    <circle cx="2" cy="8" r=".7" />
                    <circle cx="2" cy="13" r=".7" />
                  </svg>
                </button>
              </div>
              <button
                className="reset-button"
                type="button"
                onClick={resetFilters}
                disabled={!hasActiveFilters}
              >
                Reset filters
              </button>
            </div>
          </div>
        </section>

        {selectedItems.length > 0 ? (
          <section className="selection-toolbar" aria-label="Выбранные слайды">
            <div>
              <strong>Выбрано: {selectedItems.length}</strong>
              <span>Слайды будут добавлены в указанном порядке.</span>
            </div>
            <div className="selection-toolbar__actions">
              <button
                className="button button--secondary button--compact"
                type="button"
                onClick={() => setSelectedSlideIds(new Set())}
                disabled={insertingId !== null}
              >
                Отменить выбор
              </button>
              <button
                className="button button--primary button--compact"
                type="button"
                onClick={() => void insertSelectedSlides()}
                disabled={insertingId !== null}
              >
                {insertingId === MULTI_INSERT_ID ? (
                  <>
                    <span className="spinner" aria-hidden="true" />
                    Добавляем…
                  </>
                ) : (
                  <>Добавить выбранные ({selectedItems.length})</>
                )}
              </button>
            </div>
          </section>
        ) : null}

        <section
          className="catalog-content"
          aria-label="Slide catalog"
          aria-busy={activeSection === "presentations" && catalog.loading}
        >
          {activeSection === "favorites" && favoriteItems.length === 0 ? (
            <div className="state-panel">
              <span className="state-panel__icon" aria-hidden="true">
                <svg viewBox="0 0 24 24">
                  <path d="M12 20s-7-4.4-7-10a4 4 0 017-2.6A4 4 0 0119 10c0 5.6-7 10-7 10z" />
                </svg>
              </span>
              <h2>Избранное пока пусто</h2>
              <p>Добавляйте часто используемые слайды, чтобы быстро находить их здесь.</p>
              <button
                className="button button--secondary"
                type="button"
                onClick={() => setActiveSection("presentations")}
              >
                Открыть презентации
              </button>
            </div>
          ) : null}

          {activeSection === "favorites" && favoriteItems.length > 0 ? (
            <div className={`slide-grid slide-grid--${viewMode}`}>
              {favoriteItems.map((item) => (
                <SlideCard
                  key={item.id}
                  item={item}
                  previewUrl={api.getPreviewUrl(item.id)}
                  inserting={insertingId === item.id}
                  insertionBlocked={insertingId !== null}
                  favorite
                  selected={selectedSlideIds.has(item.id)}
                  onToggleFavorite={() => toggleFavorite(item.id)}
                  onToggleSelection={() => toggleSlideSelection(item.id)}
                  onOpen={() => setSelectedSlide(item)}
                  onInsert={() => void insertSlide(item)}
                />
              ))}
            </div>
          ) : null}

          {activeSection === "presentations" ? (
            <>
          {catalog.loading ? <SkeletonCatalog /> : null}

          {!catalog.loading && catalog.error ? (
            <div className="state-panel state-panel--error" role="alert">
              <span className="state-panel__icon" aria-hidden="true">
                <svg viewBox="0 0 24 24">
                  <path d="M12 3l9 16H3l9-16zM12 8v5M12 16.5v.5" />
                </svg>
              </span>
              <h2>Slide library unavailable</h2>
              <p>{catalog.error}</p>
              <button className="button button--primary" type="button" onClick={() => setRefreshKey((v) => v + 1)}>
                Retry
              </button>
            </div>
          ) : null}

          {!catalog.loading && !catalog.error && isEmptyCatalog ? (
            <div className="state-panel">
              <span className="state-panel__icon" aria-hidden="true">
                <svg viewBox="0 0 24 24">
                  <rect x="4" y="5" width="16" height="13" rx="2" />
                  <path d="M8 9h8M8 13h5" />
                </svg>
              </span>
              <h2>The library is empty</h2>
              <p>No slides are available yet. Refresh after content has been added.</p>
              <button className="button button--secondary" type="button" onClick={() => setRefreshKey((v) => v + 1)}>
                Refresh library
              </button>
            </div>
          ) : null}

          {!catalog.loading && !catalog.error && catalog.response?.items.length === 0 && !isEmptyCatalog ? (
            <div className="state-panel">
              <span className="state-panel__icon" aria-hidden="true">
                <svg viewBox="0 0 24 24">
                  <circle cx="10" cy="10" r="6" />
                  <path d="M14.5 14.5L20 20M7.5 10h5" />
                </svg>
              </span>
              <h2>No slides found</h2>
              <p>Try changing your search or filters.</p>
              <button className="button button--secondary" type="button" onClick={resetFilters}>
                Clear search and filters
              </button>
            </div>
          ) : null}

          {!catalog.loading && !catalog.error && catalog.response && catalog.response.items.length > 0 ? (
            <div className={`slide-grid slide-grid--${viewMode}`}>
              {visibleItems.map((item) => (
                <SlideCard
                  key={item.id}
                  item={item}
                  previewUrl={api.getPreviewUrl(item.id)}
                  inserting={insertingId === item.id}
                  insertionBlocked={insertingId !== null}
                  favorite={favoriteIds.has(item.id)}
                  selected={selectedSlideIds.has(item.id)}
                  onToggleFavorite={() => toggleFavorite(item.id)}
                  onToggleSelection={() => toggleSlideSelection(item.id)}
                  onOpen={() => setSelectedSlide(item)}
                  onInsert={() => void insertSlide(item)}
                />
              ))}
            </div>
          ) : null}
            </>
          ) : null}
        </section>
          </>
        ) : null}
          </div>
        </div>
      </main>

      {selectedSlide ? (
        <SlideDetailsDialog
          item={selectedSlide}
          previewUrl={api.getPreviewUrl(selectedSlide.id)}
          inserting={insertingId === selectedSlide.id}
          insertionBlocked={insertingId !== null}
          onClose={() => setSelectedSlide(null)}
          onInsert={() => void insertSlide(selectedSlide)}
        />
      ) : null}

      {toast ? <Toast message={toast} onDismiss={() => setToast(null)} /> : null}
    </div>
  );
}
