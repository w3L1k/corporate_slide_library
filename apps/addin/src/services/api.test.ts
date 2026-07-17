import { afterEach, describe, expect, it, vi } from "vitest";

import { HttpSlideLibraryApi } from "./api";

describe("HttpSlideLibraryApi", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("forwards the correct URL and method to the underlying fetch", async () => {
    const responseBody = {
      items: [],
      total: 0,
      availableCategories: []
    };

    const fetchMock = vi.fn(() =>
      Promise.resolve(
        new Response(JSON.stringify(responseBody), {
          status: 200,
          headers: { "Content-Type": "application/json" }
        })
      )
    );

    vi.stubGlobal("fetch", fetchMock);
    const api = new HttpSlideLibraryApi();

    await expect(api.listSlides({})).resolves.toEqual(responseBody);
    expect(fetchMock).toHaveBeenCalledWith("/api/slides", expect.objectContaining({ method: "GET" }));
  });

  it("downloads a personal asset only through its encoded registered ID route", async () => {
    const fetchMock = vi.fn(() =>
      Promise.resolve(
        new Response(new Uint8Array([1, 2, 3]), {
          status: 200,
          headers: { "Content-Type": "image/png" }
        })
      )
    );

    vi.stubGlobal("fetch", fetchMock);
    const api = new HttpSlideLibraryApi();

    await expect(
      api.downloadPersonalAsset("11111111-1111-4111-8111-111111111111")
    ).resolves.toEqual(new Uint8Array([1, 2, 3]).buffer);
    expect(fetchMock).toHaveBeenCalledWith(
      "/api/personal-assets/11111111-1111-4111-8111-111111111111/file",
      expect.objectContaining({
        method: "GET",
        headers: { Accept: "*/*" }
      })
    );
  });

  it("deletes a personal asset through its registered ID route", async () => {
    const fetchMock = vi.fn(() =>
      Promise.resolve(new Response(null, { status: 204 }))
    );

    vi.stubGlobal("fetch", fetchMock);
    const api = new HttpSlideLibraryApi();

    await expect(
      api.deletePersonalAsset("11111111-1111-4111-8111-111111111111")
    ).resolves.toBeUndefined();
    expect(fetchMock).toHaveBeenCalledWith(
      "/api/personal-assets/11111111-1111-4111-8111-111111111111",
      { method: "DELETE" }
    );
  });
});
