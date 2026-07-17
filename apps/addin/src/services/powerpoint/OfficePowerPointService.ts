import type { PersonalAsset } from "@slide-library/shared";
import type { SlideLibraryApi } from "../api";
import { arrayBufferToBase64 } from "./arrayBufferToBase64";
import { PowerPointUnavailableError, type PowerPointService } from "./types";

const IMAGE_INSERTION_UNAVAILABLE_MESSAGE =
  "This version of PowerPoint does not support inserting images from the personal library.";
const SVG_INSERTION_UNAVAILABLE_MESSAGE =
  "This version of PowerPoint does not support SVG insertion. Upload a PNG version of the image or update PowerPoint.";
const RASTER_IMAGE_MIME_TYPES = new Set(["image/png", "image/jpeg", "image/webp"]);

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

export type ImageCoercionType = "image" | "xmlSvg";

export interface OfficeImageInsertionRuntime {
  insert(data: string, coercionType: ImageCoercionType): Promise<void>;
}

export interface OfficePowerPointServiceOptions {
  imageInsertionRuntime?: OfficeImageInsertionRuntime;
  supportsSvgInsertion?: boolean;
}

export class OfficePowerPointService implements PowerPointService {
  private readonly api: Pick<SlideLibraryApi, "downloadSlide" | "downloadPersonalAsset">;
  private readonly runtime: PowerPointRuntime;
  private readonly imageInsertionRuntime: OfficeImageInsertionRuntime | undefined;
  private readonly supportsSvgInsertion: boolean;

  constructor(
    api: Pick<SlideLibraryApi, "downloadSlide" | "downloadPersonalAsset">,
    runtime: PowerPointRuntime,
    options: OfficePowerPointServiceOptions = {}
  ) {
    this.api = api;
    this.runtime = runtime;
    this.imageInsertionRuntime = options.imageInsertionRuntime;
    this.supportsSvgInsertion = options.supportsSvgInsertion ?? false;
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
    await this.insertPresentationFiles(slideFiles);
  }

  async insertPersonalAsset(asset: PersonalAsset): Promise<void> {
    const file = await this.api.downloadPersonalAsset(asset.id);

    if (asset.kind === "presentation") {
      await this.insertPresentationFiles([file]);
      return;
    }

    if (asset.mimeType === "image/svg+xml") {
      if (!this.imageInsertionRuntime || !this.supportsSvgInsertion) {
        throw new PowerPointUnavailableError(SVG_INSERTION_UNAVAILABLE_MESSAGE);
      }

      await this.imageInsertionRuntime.insert(new TextDecoder().decode(file), "xmlSvg");
      return;
    }

    if (!RASTER_IMAGE_MIME_TYPES.has(asset.mimeType)) {
      throw new Error(`Unsupported personal image format: ${asset.mimeType}`);
    }
    if (!this.imageInsertionRuntime) {
      throw new PowerPointUnavailableError(IMAGE_INSERTION_UNAVAILABLE_MESSAGE);
    }

    await this.imageInsertionRuntime.insert(arrayBufferToBase64(file), "image");
  }

  private async insertPresentationFiles(slideFiles: readonly ArrayBuffer[]): Promise<void> {
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
