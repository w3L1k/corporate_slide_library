import { describe, expect, it, vi } from "vitest";
import type { SlideLibraryApi } from "../api";
import {
  OfficePowerPointService,
  type PowerPointRequestContextRuntime,
  type PowerPointRuntime
} from "./OfficePowerPointService";

describe("OfficePowerPointService", () => {
  it("downloads the registered PPTX and inserts it through a synced PowerPoint.run batch", async () => {
    const events: string[] = [];
    const insertSlidesFromBase64 = vi.fn(() => events.push("insert"));
    const sync = vi.fn(async () => {
      events.push("sync");
    });
    const context: PowerPointRequestContextRuntime = {
      presentation: { insertSlidesFromBase64 },
      sync
    };
    let runCalls = 0;
    const runtime: PowerPointRuntime = {
      async run<T>(batch: (value: PowerPointRequestContextRuntime) => Promise<T>): Promise<T> {
        runCalls += 1;
        events.push("run");
        return batch(context);
      }
    };
    const downloadSlide = vi.fn(async () => new Uint8Array([1, 2, 3]).buffer);
    const api: Pick<SlideLibraryApi, "downloadSlide"> = { downloadSlide };
    const service = new OfficePowerPointService(api, runtime);

    await service.insertSlide("revenue-overview");

    expect(service.isAvailable()).toBe(true);
    expect(service.getUnavailableReason()).toBeNull();
    expect(downloadSlide).toHaveBeenCalledWith("revenue-overview");
    expect(runCalls).toBe(1);
    expect(insertSlidesFromBase64).toHaveBeenCalledWith("AQID", {
      formatting: "KeepSourceFormatting"
    });
    expect(sync).toHaveBeenCalledTimes(1);
    expect(events).toEqual(["run", "insert", "sync"]);
  });

  it("downloads and inserts multiple slides in catalog order with one sync", async () => {
    const insertSlidesFromBase64 = vi.fn();
    const sync = vi.fn(async () => undefined);
    const context: PowerPointRequestContextRuntime = {
      presentation: { insertSlidesFromBase64 },
      sync
    };
    const runtime: PowerPointRuntime = {
      run: async <T>(batch: (value: PowerPointRequestContextRuntime) => Promise<T>) =>
        batch(context)
    };
    const downloadSlide = vi.fn(async (id: string) =>
      new TextEncoder().encode(id).buffer
    );
    const service = new OfficePowerPointService({ downloadSlide }, runtime);

    await service.insertSlides(["first-slide", "second-slide"]);

    expect(downloadSlide.mock.calls.map(([id]) => id)).toEqual([
      "first-slide",
      "second-slide"
    ]);
    expect(insertSlidesFromBase64).toHaveBeenCalledTimes(2);
    expect(sync).toHaveBeenCalledTimes(1);
  });
});
