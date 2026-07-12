import type { SlideLibraryItem } from "@slide-library/shared";

export interface SlideStorage {
  getCatalog(): Promise<SlideLibraryItem[]>;
  getItem(id: string): Promise<SlideLibraryItem | undefined>;
  getSlide(id: string): Promise<Buffer>;
  getPreview(id: string): Promise<Buffer>;
  refresh(): Promise<SlideLibraryItem[]>;
}
