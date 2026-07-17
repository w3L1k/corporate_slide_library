import type { SlideLibraryItem, SlideListResponse } from "@slide-library/shared";
import { vi } from "vitest";
import type { SlideLibraryApi } from "../services/api";
import type { PowerPointService } from "../services/powerpoint";

export const revenueSlide: SlideLibraryItem = {
  id: "revenue-overview",
  title: "Revenue overview",
  description: "Quarterly revenue performance with an approved executive summary.",
  category: "Finance",
  tags: ["revenue", "quarterly", "executive"],
  department: "Finance",
  language: "en-US",
  version: "2.4",
  status: "approved",
  updatedAt: "2026-06-18T10:30:00.000Z",
  sourceFile: "slides/revenue-overview.pptx",
  previewFile: "previews/revenue-overview.png",
  author: "Alex Morgan",
  owner: "Finance Operations"
};

export const strategySlide: SlideLibraryItem = {
  ...revenueSlide,
  id: "strategy-roadmap",
  title: "Strategy roadmap",
  description: "Product strategy milestones and delivery horizons.",
  category: "Strategy",
  tags: ["strategy", "roadmap"],
  department: "Strategy",
  version: "1.3",
  updatedAt: "2026-05-02T08:00:00.000Z",
  sourceFile: "slides/strategy-roadmap.pptx",
  previewFile: "previews/strategy-roadmap.png",
  owner: "Strategy Office"
};

export const catalogResponse: SlideListResponse = {
  items: [revenueSlide],
  total: 1,
  availableCategories: ["Finance", "Strategy"]
};

export const createApi = (
  listSlidesImplementation: SlideLibraryApi["listSlides"] = async () => catalogResponse
): SlideLibraryApi => ({
  listSlides: vi.fn(listSlidesImplementation),
  getSlide: vi.fn(async () => revenueSlide),
  downloadSlide: vi.fn(async () => new ArrayBuffer(0)),
  getPreviewUrl: vi.fn((id: string) => `/api/slides/${encodeURIComponent(id)}/preview`),
  listPersonalAssets: vi.fn(async () => ({ items: [], total: 0 })),
  uploadPersonalAsset: vi.fn(async (kind, title, file) => ({
    id: "11111111-1111-4111-8111-111111111111",
    title,
    kind,
    fileName: file.name,
    mimeType: file.type || "application/octet-stream",
    size: file.size,
    createdAt: "2026-07-17T10:00:00.000Z"
  })),
  downloadPersonalAsset: vi.fn(async () => new ArrayBuffer(0)),
  getPersonalAssetFileUrl: vi.fn(
    (id: string) => `/api/personal-assets/${encodeURIComponent(id)}/file`
  )
});

export const createAvailablePowerPointService = (
  insertSlideImplementation: PowerPointService["insertSlide"] = async () => undefined,
  insertPersonalAssetImplementation: PowerPointService["insertPersonalAsset"] = async () =>
    undefined
): PowerPointService => ({
  isAvailable: () => true,
  getUnavailableReason: () => null,
  insertSlide: vi.fn(insertSlideImplementation),
  insertSlides: vi.fn(async (slideIds: readonly string[]) => {
    for (const slideId of slideIds) {
      await insertSlideImplementation(slideId);
    }
  }),
  insertPersonalAsset: vi.fn(insertPersonalAssetImplementation)
});

export interface Deferred<T> {
  promise: Promise<T>;
  resolve(value: T): void;
  reject(reason: unknown): void;
}

export const createDeferred = <T>(): Deferred<T> => {
  let resolvePromise: (value: T | PromiseLike<T>) => void = () => undefined;
  let rejectPromise: (reason?: unknown) => void = () => undefined;
  const promise = new Promise<T>((resolve, reject) => {
    resolvePromise = resolve;
    rejectPromise = reject;
  });

  return {
    promise,
    resolve: (value) => resolvePromise(value),
    reject: (reason) => rejectPromise(reason)
  };
};
