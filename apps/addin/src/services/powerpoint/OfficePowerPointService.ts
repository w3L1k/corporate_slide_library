import type { SlideLibraryApi } from "../api";
import { arrayBufferToBase64 } from "./arrayBufferToBase64";
import type { PowerPointService } from "./types";

export interface PresentationRuntime {
  insertSlidesFromBase64(
    base64File: string,
    options?: { formatting?: "KeepSourceFormatting" | "UseDestinationTheme" }
  ): void;
}

export interface PowerPointRequestContextRuntime {
  presentation: PresentationRuntime;
  sync(): Promise<void>;
}

export interface PowerPointRuntime {
  run<T>(batch: (context: PowerPointRequestContextRuntime) => Promise<T>): Promise<T>;
}

export class OfficePowerPointService implements PowerPointService {
  private readonly api: Pick<SlideLibraryApi, "downloadSlide">;
  private readonly runtime: PowerPointRuntime;

  constructor(api: Pick<SlideLibraryApi, "downloadSlide">, runtime: PowerPointRuntime) {
    this.api = api;
    this.runtime = runtime;
  }

  isAvailable(): boolean {
    return true;
  }

  getUnavailableReason(): null {
    return null;
  }

  async insertSlide(slideId: string): Promise<void> {
    await this.insertSlides([slideId]);
  }

  async insertSlides(slideIds: readonly string[]): Promise<void> {
    if (slideIds.length === 0) {
      return;
    }

    const slideFiles = await Promise.all(slideIds.map((slideId) => this.api.downloadSlide(slideId)));
    const base64Files = slideFiles.map(arrayBufferToBase64);

    await this.runtime.run(async (context) => {
      for (const base64File of base64Files) {
        context.presentation.insertSlidesFromBase64(base64File, {
          formatting: "KeepSourceFormatting"
        });
      }
      await context.sync();
    });
  }
}
