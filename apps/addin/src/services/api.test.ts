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
});
