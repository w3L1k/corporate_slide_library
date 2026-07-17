import { describe, expect, it, vi } from "vitest";
import type { PersonalAsset } from "@slide-library/shared";
import type { SlideLibraryApi } from "../api";
import {
  OfficePowerPointService,
  type PowerPointRequestContextRuntime,
  type PowerPointRuntime
} from "./OfficePowerPointService";
import { PowerPointUnavailableError } from "./types";

const photoAsset: PersonalAsset = {
  id: "11111111-1111-4111-8111-111111111111",
  title: "Team photo",
  kind: "photo",
  fileName: "team.png",
  mimeType: "image/png",
  size: 3,
  createdAt: "2026-07-17T10:00:00.000Z"
};

const unusedPowerPointRuntime: PowerPointRuntime = {
  run<T>(): Promise<T> {
    return Promise.reject(new Error("PowerPoint.run was not expected in this test."));
  }
};

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
    const downloadPersonalAsset = vi.fn(async () => new ArrayBuffer(0));
    const api: Pick<SlideLibraryApi, "downloadSlide" | "downloadPersonalAsset"> = {
      downloadSlide,
      downloadPersonalAsset
    };
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
    const downloadPersonalAsset = vi.fn(async () => new ArrayBuffer(0));
    const service = new OfficePowerPointService(
      { downloadSlide, downloadPersonalAsset },
      runtime
    );

    await service.insertSlides(["first-slide", "second-slide"]);

    expect(downloadSlide.mock.calls.map(([id]) => id)).toEqual([
      "first-slide",
      "second-slide"
    ]);
    expect(insertSlidesFromBase64).toHaveBeenCalledTimes(2);
    expect(sync).toHaveBeenCalledTimes(1);
  });

  it("inserts a personal raster image through Office image coercion", async () => {
    const downloadPersonalAsset = vi.fn(async () => new Uint8Array([1, 2, 3]).buffer);
    const insert = vi.fn(async () => undefined);
    const service = new OfficePowerPointService(
      {
        downloadSlide: vi.fn(async () => new ArrayBuffer(0)),
        downloadPersonalAsset
      },
      unusedPowerPointRuntime,
      { imageInsertionRuntime: { insert } }
    );

    await service.insertPersonalAsset(photoAsset);

    expect(downloadPersonalAsset).toHaveBeenCalledWith(photoAsset.id);
    expect(insert).toHaveBeenCalledWith("AQID", "image");
  });

  it("inserts a safe personal SVG as XML when ImageCoercion 1.2 is available", async () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg"><path d="M0 0"/></svg>';
    const insert = vi.fn(async () => undefined);
    const service = new OfficePowerPointService(
      {
        downloadSlide: vi.fn(async () => new ArrayBuffer(0)),
        downloadPersonalAsset: vi.fn(async () => new TextEncoder().encode(svg).buffer)
      },
      unusedPowerPointRuntime,
      {
        imageInsertionRuntime: { insert },
        supportsSvgInsertion: true
      }
    );

    await service.insertPersonalAsset({
      ...photoAsset,
      id: "22222222-2222-4222-8222-222222222222",
      kind: "logo",
      fileName: "brand.svg",
      mimeType: "image/svg+xml"
    });

    expect(insert).toHaveBeenCalledWith(svg, "xmlSvg");
  });

  it("inserts a personal PPTX with source formatting", async () => {
    const insertSlidesFromBase64 = vi.fn();
    const sync = vi.fn(async () => undefined);
    const runtime: PowerPointRuntime = {
      run: async <T>(batch: (context: PowerPointRequestContextRuntime) => Promise<T>) =>
        batch({
          presentation: { insertSlidesFromBase64 },
          sync
        })
    };
    const service = new OfficePowerPointService(
      {
        downloadSlide: vi.fn(async () => new ArrayBuffer(0)),
        downloadPersonalAsset: vi.fn(async () => new Uint8Array([1, 2, 3]).buffer)
      },
      runtime
    );

    await service.insertPersonalAsset({
      ...photoAsset,
      id: "33333333-3333-4333-8333-333333333333",
      kind: "presentation",
      fileName: "personal.pptx",
      mimeType:
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    });

    expect(insertSlidesFromBase64).toHaveBeenCalledWith("AQID", {
      formatting: "KeepSourceFormatting"
    });
    expect(sync).toHaveBeenCalledTimes(1);
  });

  it("reports a friendly compatibility error when image coercion is unavailable", async () => {
    const service = new OfficePowerPointService(
      {
        downloadSlide: vi.fn(async () => new ArrayBuffer(0)),
        downloadPersonalAsset: vi.fn(async () => new Uint8Array([1, 2, 3]).buffer)
      },
      unusedPowerPointRuntime
    );

    await expect(service.insertPersonalAsset(photoAsset)).rejects.toBeInstanceOf(
      PowerPointUnavailableError
    );
  });
});
