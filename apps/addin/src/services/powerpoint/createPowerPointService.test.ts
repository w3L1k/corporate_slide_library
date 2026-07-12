import { afterEach, describe, expect, it, vi } from "vitest";
import type { SlideLibraryApi } from "../api";
import { BrowserPowerPointService } from "./BrowserPowerPointService";
import { createPowerPointService, type OfficeRuntime } from "./createPowerPointService";
import { OfficePowerPointService, type PowerPointRuntime } from "./OfficePowerPointService";
import { BROWSER_POWERPOINT_MESSAGE, PowerPointUnavailableError } from "./types";

const api: Pick<SlideLibraryApi, "downloadSlide"> = {
  downloadSlide: vi.fn(async () => new ArrayBuffer(0))
};

const unusedPowerPointRuntime: PowerPointRuntime = {
  run<T>(): Promise<T> {
    return Promise.reject(new Error("PowerPoint.run was not expected in this test."));
  }
};

const createOfficeRuntime = (host: string, supported: boolean): OfficeRuntime => ({
  onReady: vi.fn(async () => ({ host })),
  context: {
    requirements: {
      isSetSupported: vi.fn(() => supported)
    }
  }
});

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

describe("createPowerPointService", () => {
  it("returns immediately in browser mode when Office-specific globals are absent", async () => {
    const service = await createPowerPointService(api, {});

    expect(service).toBeInstanceOf(BrowserPowerPointService);
    expect(service.isAvailable()).toBe(false);
    expect(service.getUnavailableReason()).toBe(BROWSER_POWERPOINT_MESSAGE);
    await expect(service.insertSlide("revenue-overview")).rejects.toBeInstanceOf(
      PowerPointUnavailableError
    );
  });

  it("waits for Office.js before resolving a late PowerPoint namespace", async () => {
    const powerPoint = { ...unusedPowerPointRuntime };
    const office = createOfficeRuntime("PowerPoint", true);
    vi.stubGlobal("Office", office);
    vi.stubGlobal("PowerPoint", undefined);

    const servicePromise = createPowerPointService(api);
    vi.stubGlobal("PowerPoint", powerPoint);

    expect(await servicePromise).toBeInstanceOf(OfficePowerPointService);
  });

  it("does not hang if Office.onReady never settles", async () => {
    vi.useFakeTimers();
    const office: OfficeRuntime = {
      onReady: () => new Promise<never>(() => undefined),
      context: {
        requirements: {
          isSetSupported: () => true
        }
      }
    };

    const servicePromise = createPowerPointService(
      api,
      { office, powerPoint: unusedPowerPointRuntime },
      50
    );
    await vi.advanceTimersByTimeAsync(50);
    const service = await servicePromise;

    expect(service).toBeInstanceOf(BrowserPowerPointService);
    expect(service.getUnavailableReason()).toBe(BROWSER_POWERPOINT_MESSAGE);
  });

  it("uses browser fallback when Office is ready in a different host", async () => {
    const office = createOfficeRuntime("Excel", true);

    const service = await createPowerPointService(api, {
      office,
      powerPoint: unusedPowerPointRuntime
    });

    expect(office.onReady).toHaveBeenCalledTimes(1);
    expect(service).toBeInstanceOf(BrowserPowerPointService);
    expect(service.getUnavailableReason()).toBe(BROWSER_POWERPOINT_MESSAGE);
  });

  it("reports a friendly incompatibility when PowerPointApi 1.2 is unavailable", async () => {
    const office = createOfficeRuntime("PowerPoint", false);

    const service = await createPowerPointService(api, {
      office,
      powerPoint: unusedPowerPointRuntime
    });

    expect(office.context?.requirements?.isSetSupported).toHaveBeenCalledWith(
      "PowerPointApi",
      "1.2"
    );
    expect(service).toBeInstanceOf(BrowserPowerPointService);
    expect(service.getUnavailableReason()).toMatch(/does not support inserting slides/i);
  });

  it("creates the real Office service only for a supported PowerPoint host", async () => {
    const office = createOfficeRuntime("PowerPoint", true);

    const service = await createPowerPointService(api, {
      office,
      powerPoint: unusedPowerPointRuntime
    });

    expect(service).toBeInstanceOf(OfficePowerPointService);
    expect(service.isAvailable()).toBe(true);
    expect(service.getUnavailableReason()).toBeNull();
  });

  it("treats missing requirement detection as incompatible", async () => {
    const office: OfficeRuntime = {
      onReady: vi.fn(async () => ({ host: "PowerPoint" }))
    };

    const service = await createPowerPointService(api, {
      office,
      powerPoint: unusedPowerPointRuntime
    });

    expect(service).toBeInstanceOf(BrowserPowerPointService);
    expect(service.getUnavailableReason()).toMatch(/update PowerPoint/i);
  });
});
