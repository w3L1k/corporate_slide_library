import type { PersonalAsset } from "@slide-library/shared";

export const BROWSER_POWERPOINT_MESSAGE =
  "PowerPoint integration is available when the add-in is running inside Microsoft PowerPoint.";

export interface PowerPointService {
  isAvailable(): boolean;
  getUnavailableReason(): string | null;
  insertSlide(slideId: string): Promise<void>;
  insertSlides(slideIds: readonly string[]): Promise<void>;
  insertPersonalAsset(asset: PersonalAsset): Promise<void>;
}

export class PowerPointUnavailableError extends Error {
  constructor(message = BROWSER_POWERPOINT_MESSAGE) {
    super(message);
    this.name = "PowerPointUnavailableError";
  }
}
