import {
  BROWSER_POWERPOINT_MESSAGE,
  PowerPointUnavailableError,
  type PowerPointService
} from "./types";

export class BrowserPowerPointService implements PowerPointService {
  private readonly unavailableMessage: string;

  constructor(unavailableMessage = BROWSER_POWERPOINT_MESSAGE) {
    this.unavailableMessage = unavailableMessage;
  }

  isAvailable(): boolean {
    return false;
  }

  getUnavailableReason(): string {
    return this.unavailableMessage;
  }

  async insertSlide(slideId: string): Promise<void> {
    void slideId;
    throw new PowerPointUnavailableError(this.unavailableMessage);
  }

  async insertSlides(slideIds: readonly string[]): Promise<void> {
    void slideIds;
    throw new PowerPointUnavailableError(this.unavailableMessage);
  }
}
