import { useCallback, useEffect, useMemo, useState } from "react";
import type { SlideLibraryItem, SlideListResponse, SlideStatus } from "@slide-library/shared";
import { SkeletonCatalog } from "./components/SkeletonCatalog";
import { SlideCard } from "./components/SlideCard";
import { SlideDetailsDialog } from "./components/SlideDetailsDialog";
import { Toast, type ToastMessage } from "./components/Toast";
import { useDebouncedValue } from "./hooks/useDebouncedValue";
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

const DEFAULT_STATUS: StatusFilter = "approved";
const DEFAULT_DEBOUNCE_MS = 300;

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
  const debouncedQuery = useDebouncedValue(query.trim(), searchDebounceMs);
  const [category, setCategory] = useState("");
  const [status, setStatus] = useState<StatusFilter>(DEFAULT_STATUS);
  const [refreshKey, setRefreshKey] = useState(0);
  const [catalog, setCatalog] = useState<CatalogState>(initialCatalogState);
  const [categories, setCategories] = useState<string[]>([]);
  const [selectedSlide, setSelectedSlide] = useState<SlideLibraryItem | null>(null);
  const [insertingId, setInsertingId] = useState<string | null>(null);
  const [toast, setToast] = useState<ToastMessage | null>(null);
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

  return (
    <div className="app-shell">
      <header className="app-header">
        <div className="brand">
          <span className="brand__mark" aria-hidden="true">
            <svg viewBox="0 0 24 24">
              <rect x="4" y="5" width="16" height="11" rx="2" />
              <path d="M8 9h5M8 12h8M9 20h6M12 16v4" />
            </svg>
          </span>
          <div>
            <h1>Slide Library</h1>
            <p>{currentSubtitle}</p>
          </div>
        </div>
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
      </header>

      <main>
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
            <button className="reset-button" type="button" onClick={resetFilters} disabled={!hasActiveFilters}>
              Reset filters
            </button>
          </div>
        </section>

        <section className="catalog-content" aria-label="Slide catalog" aria-busy={catalog.loading}>
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
            <div className="slide-grid">
              {catalog.response.items.map((item) => (
                <SlideCard
                  key={item.id}
                  item={item}
                  previewUrl={api.getPreviewUrl(item.id)}
                  inserting={insertingId === item.id}
                  insertionBlocked={insertingId !== null}
                  onOpen={() => setSelectedSlide(item)}
                  onInsert={() => void insertSlide(item)}
                />
              ))}
            </div>
          ) : null}
        </section>
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
