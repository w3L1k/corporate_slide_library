import type { SlideLibraryItem } from "@slide-library/shared";
import { mkdir, mkdtemp, readdir, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import request from "supertest";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { createApp } from "../src/app.js";
import type { Logger } from "../src/logger.js";
import { FileSystemSlideStorage } from "../src/storage/FileSystemSlideStorage.js";
import { PersonalAssetStorage } from "../src/storage/PersonalAssetStorage.js";
import { resolveContainedPath } from "../src/storage/pathSafety.js";

const silentLogger: Logger = {
  info: () => undefined,
  warn: () => undefined,
  error: () => undefined
};

const catalog: SlideLibraryItem[] = [
  {
    id: "revenue-overview",
    title: "Revenue Overview",
    description: "Quarterly growth and recurring revenue",
    category: "Finance",
    tags: ["ARR", "Growth"],
    version: "2.1",
    status: "approved",
    updatedAt: "2026-07-01T10:00:00.000Z",
    sourceFile: "slides/revenue-overview.pptx",
    previewFile: "previews/revenue-overview.png",
    searchText: "board metrics"
  },
  {
    id: "customer-journey",
    title: "Customer Journey",
    description: "Lifecycle touchpoints",
    category: "Marketing",
    tags: ["Lifecycle", "Experience"],
    version: "1.0",
    status: "draft",
    updatedAt: "2026-06-15T09:00:00.000Z",
    sourceFile: "slides/customer-journey.pptx",
    previewFile: "previews/customer-journey.jpg"
  },
  {
    id: "missing-asset",
    title: "Missing Asset Example",
    category: "Finance",
    tags: ["Missing"],
    version: "1.0",
    status: "deprecated",
    updatedAt: "2026-05-01T08:00:00.000Z",
    sourceFile: "slides/missing-asset.pptx",
    previewFile: "previews/missing-asset.png"
  }
];

describe("slide catalog API", () => {
  let libraryPath: string;
  let app: ReturnType<typeof createApp>;

  beforeEach(async () => {
    libraryPath = await mkdtemp(path.join(os.tmpdir(), "slide-library-server-"));
    await Promise.all([
      mkdir(path.join(libraryPath, "slides")),
      mkdir(path.join(libraryPath, "previews"))
    ]);
    await Promise.all([
      writeFile(
        path.join(libraryPath, "catalog.json"),
        JSON.stringify(catalog, null, 2),
        "utf8"
      ),
      writeFile(path.join(libraryPath, "slides", "revenue-overview.pptx"), "pptx-revenue"),
      writeFile(path.join(libraryPath, "previews", "revenue-overview.png"), "png-revenue"),
      writeFile(path.join(libraryPath, "slides", "customer-journey.pptx"), "pptx-journey"),
      writeFile(path.join(libraryPath, "previews", "customer-journey.jpg"), "jpg-journey")
    ]);

    const storage = new FileSystemSlideStorage(libraryPath, { logger: silentLogger });
    const personalAssetStorage = new PersonalAssetStorage(libraryPath);
    await personalAssetStorage.initialize();
    app = createApp({
      storage,
      personalAssetStorage,
      config: { corsOrigins: ["http://localhost:3000"], enableAdminReindex: false },
      logger: silentLogger
    });
  });

  afterEach(async () => {
    await rm(libraryPath, { recursive: true, force: true });
  });

  it("returns health and the complete catalog", async () => {
    await request(app).get("/api/health").expect(200, { status: "ok" });

    const response = await request(app).get("/api/slides").expect(200);
    expect(response.body.total).toBe(3);
    expect(response.body.items.map((item: SlideLibraryItem) => item.id)).toEqual([
      "revenue-overview",
      "customer-journey",
      "missing-asset"
    ]);
    expect(response.body.availableCategories).toEqual(["Finance", "Marketing"]);
  });

  it("returns a valid empty response for an empty catalog", async () => {
    await writeFile(path.join(libraryPath, "catalog.json"), "[]", "utf8");

    const response = await request(app).get("/api/slides").expect(200);
    expect(response.body).toEqual({ items: [], total: 0, availableCategories: [] });
  });

  it.each([
    ["REVENUE", "revenue-overview"],
    ["recurring", "revenue-overview"],
    ["arr", "revenue-overview"],
    ["finance", "revenue-overview"],
    ["BOARD METRICS", "revenue-overview"],
    ["lifecycle", "customer-journey"]
  ])("searches all supported fields case-insensitively: %s", async (query, expectedId) => {
    const response = await request(app).get("/api/slides").query({ q: query }).expect(200);
    expect(response.body.items.map((item: SlideLibraryItem) => item.id)).toContain(expectedId);
  });

  it("combines case-insensitive category and status filters", async () => {
    const response = await request(app)
      .get("/api/slides")
      .query({ category: "fINAnce", status: "APPROVED" })
      .expect(200);

    expect(response.body.total).toBe(1);
    expect(response.body.items[0].id).toBe("revenue-overview");
  });

  it("filters by category independently", async () => {
    const response = await request(app)
      .get("/api/slides")
      .query({ category: "finance" })
      .expect(200);

    expect(response.body.items.map((item: SlideLibraryItem) => item.id)).toEqual([
      "revenue-overview",
      "missing-asset"
    ]);
  });

  it("filters by status independently", async () => {
    const response = await request(app)
      .get("/api/slides")
      .query({ status: "draft" })
      .expect(200);

    expect(response.body.items.map((item: SlideLibraryItem) => item.id)).toEqual([
      "customer-journey"
    ]);
  });

  it("combines search, category, and status as an AND query", async () => {
    const response = await request(app)
      .get("/api/slides")
      .query({ q: "growth", category: "Finance", status: "approved" })
      .expect(200);

    expect(response.body.items.map((item: SlideLibraryItem) => item.id)).toEqual([
      "revenue-overview"
    ]);
  });

  it("rejects an unknown status filter", async () => {
    const response = await request(app)
      .get("/api/slides")
      .query({ status: "published" })
      .expect(400);
    expect(response.body.error.code).toBe("INVALID_STATUS");
  });

  it("returns one catalog item and a friendly error for an unknown id", async () => {
    const itemResponse = await request(app).get("/api/slides/revenue-overview").expect(200);
    expect(itemResponse.body.title).toBe("Revenue Overview");

    const missingResponse = await request(app).get("/api/slides/not-registered").expect(404);
    expect(missingResponse.body.error.code).toBe("SLIDE_NOT_FOUND");
  });

  it("does not expose binary routes for an unregistered id", async () => {
    const fileResponse = await request(app).get("/api/slides/not-registered/file").expect(404);
    expect(fileResponse.body.error.code).toBe("SLIDE_NOT_FOUND");

    const previewResponse = await request(app)
      .get("/api/slides/not-registered/preview")
      .expect(404);
    expect(previewResponse.body.error.code).toBe("SLIDE_NOT_FOUND");
  });

  it("serves only the registered PowerPoint file", async () => {
    const response = await request(app)
      .get("/api/slides/revenue-overview/file")
      .buffer(true)
      .parse((incoming, complete) => {
        const chunks: Buffer[] = [];
        incoming.on("data", (chunk: Buffer | string) => chunks.push(Buffer.from(chunk)));
        incoming.on("end", () => complete(null, Buffer.concat(chunks)));
      })
      .expect(200);
    expect(response.headers["content-type"]).toContain(
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    expect(response.headers["content-disposition"]).toBe(
      'attachment; filename="revenue-overview.pptx"'
    );
    expect(Buffer.isBuffer(response.body)).toBe(true);
    expect(response.body.toString()).toBe("pptx-revenue");
  });

  it("serves previews with their registered image type", async () => {
    const response = await request(app).get("/api/slides/customer-journey/preview").expect(200);
    expect(response.headers["content-type"]).toContain("image/jpeg");
  });

  it("uploads and serves registered personal presentations, photos, and logos", async () => {
    const photo = await request(app)
      .post("/api/personal-assets")
      .field("kind", "photo")
      .field("title", "Командная фотография")
      .attach(
        "file",
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1]),
        "team.png"
      )
      .expect(201);
    const logo = await request(app)
      .post("/api/personal-assets")
      .field("kind", "logo")
      .field("title", "Логотип продукта")
      .attach("file", Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"></svg>'), "logo.svg")
      .expect(201);
    const presentation = await request(app)
      .post("/api/personal-assets")
      .field("kind", "presentation")
      .field("title", "Личная презентация")
      .attach("file", Buffer.from("PK\u0003\u0004pptx"), "personal.pptx")
      .expect(201);

    expect(photo.body).toMatchObject({ kind: "photo", mimeType: "image/png" });
    expect(logo.body).toMatchObject({ kind: "logo", mimeType: "image/svg+xml" });
    expect(presentation.body).toMatchObject({
      kind: "presentation",
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    });

    const list = await request(app).get("/api/personal-assets").expect(200);
    expect(list.body.total).toBe(3);
    expect(list.body.items.map((item: { kind: string }) => item.kind).sort()).toEqual([
      "logo",
      "photo",
      "presentation"
    ]);

    const file = await request(app)
      .get(`/api/personal-assets/${photo.body.id}/file`)
      .expect("Content-Type", /image\/png/)
      .expect(200);
    expect(Buffer.from(file.body).subarray(0, 8)).toEqual(
      Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    );

    await request(app).delete(`/api/personal-assets/${photo.body.id}`).expect(204);
    const listAfterDelete = await request(app).get("/api/personal-assets").expect(200);
    expect(listAfterDelete.body.total).toBe(2);
    expect(
      listAfterDelete.body.items.map((item: { id: string }) => item.id)
    ).not.toContain(photo.body.id);
    await request(app)
      .get(`/api/personal-assets/${photo.body.id}/file`)
      .expect(404);
    expect(
      await readdir(path.join(libraryPath, "personal-assets", "files"))
    ).not.toContainEqual(expect.stringContaining(photo.body.id));
  });

  it("rejects invalid or unregistered personal asset deletion IDs", async () => {
    const invalid = await request(app)
      .delete("/api/personal-assets/not-a-uuid")
      .expect(400);
    expect(invalid.body.error.code).toBe("INVALID_PERSONAL_ASSET_ID");

    const missing = await request(app)
      .delete("/api/personal-assets/11111111-1111-4111-8111-111111111111")
      .expect(404);
    expect(missing.body.error.code).toBe("PERSONAL_ASSET_NOT_FOUND");
  });

  it("rejects unsupported or unsafe personal uploads", async () => {
    const unsupported = await request(app)
      .post("/api/personal-assets")
      .field("kind", "photo")
      .field("title", "Not an image")
      .attach("file", Buffer.from("plain text"), "notes.txt")
      .expect(400);
    expect(unsupported.body.error.code).toBe("UNSUPPORTED_PERSONAL_ASSET");

    const unsafeSvg = await request(app)
      .post("/api/personal-assets")
      .field("kind", "logo")
      .field("title", "Unsafe logo")
      .attach("file", Buffer.from("<svg><script>alert(1)</script></svg>"), "unsafe.svg")
      .expect(400);
    expect(unsafeSvg.body.error.code).toBe("UNSUPPORTED_PERSONAL_ASSET");
  });

  it("returns friendly errors when registered files are missing", async () => {
    const fileResponse = await request(app).get("/api/slides/missing-asset/file").expect(404);
    expect(fileResponse.body.error.code).toBe("SLIDE_FILE_NOT_FOUND");

    const previewResponse = await request(app)
      .get("/api/slides/missing-asset/preview")
      .expect(404);
    expect(previewResponse.body.error.code).toBe("PREVIEW_NOT_FOUND");
  });

  it("distinguishes an invalid catalog from an unavailable catalog", async () => {
    await writeFile(path.join(libraryPath, "catalog.json"), "{not-json", "utf8");
    const invalidResponse = await request(app).get("/api/slides").expect(500);
    expect(invalidResponse.body.error.code).toBe("CATALOG_INVALID");

    await rm(path.join(libraryPath, "catalog.json"));
    const unavailableResponse = await request(app).get("/api/slides").expect(503);
    expect(unavailableResponse.body.error.code).toBe("CATALOG_UNAVAILABLE");
  });

  it("blocks path traversal and never treats a path as an id", async () => {
    const response = await request(app)
      .get("/api/slides/%2e%2e%2fpackage.json/file")
      .expect(400);
    expect(response.body.error.code).toBe("INVALID_SLIDE_ID");
    expect(() => resolveContainedPath(libraryPath, "../package.json")).toThrow(/unsafe/i);
  });

  it("does not expose the pilot reindex route unless explicitly enabled", async () => {
    const response = await request(app).post("/api/admin/reindex").expect(404);
    expect(response.body.error.code).toBe("API_ROUTE_NOT_FOUND");
  });

  it("returns client errors for malformed or oversized JSON bodies", async () => {
    const malformedResponse = await request(app)
      .post("/api/admin/reindex")
      .set("Content-Type", "application/json")
      .send('{"invalid"')
      .expect(400);
    expect(malformedResponse.body.error.code).toBe("INVALID_JSON");

    const oversizedResponse = await request(app)
      .post("/api/admin/reindex")
      .set("Content-Type", "application/json")
      .send(JSON.stringify({ value: "x".repeat(17 * 1024) }))
      .expect(413);
    expect(oversizedResponse.body.error.code).toBe("REQUEST_TOO_LARGE");
  });

  it("refreshes through the pilot route only when explicitly enabled", async () => {
    const enabledApp = createApp({
      storage: new FileSystemSlideStorage(libraryPath, { logger: silentLogger }),
      config: { corsOrigins: [], enableAdminReindex: true },
      logger: silentLogger
    });
    const response = await request(enabledApp).post("/api/admin/reindex").expect(200);
    expect(response.body).toMatchObject({ status: "ok", itemCount: 3 });
    expect(Number.isNaN(Date.parse(response.body.refreshedAt))).toBe(false);
  });

  it("enforces the configured CORS allowlist", async () => {
    await request(app)
      .get("/api/health")
      .set("Origin", "http://localhost:3000")
      .expect("Access-Control-Allow-Origin", "http://localhost:3000")
      .expect(200);

    const response = await request(app)
      .get("/api/health")
      .set("Origin", "https://untrusted.example")
      .expect(403);
    expect(response.body.error.code).toBe("CORS_NOT_ALLOWED");
  });

  it("handles allowed and denied CORS preflight requests", async () => {
    await request(app)
      .options("/api/personal-assets/11111111-1111-4111-8111-111111111111")
      .set("Origin", "http://localhost:3000")
      .set("Access-Control-Request-Method", "DELETE")
      .expect("Access-Control-Allow-Origin", "http://localhost:3000")
      .expect("Access-Control-Allow-Methods", /DELETE/)
      .expect(204);

    const deniedResponse = await request(app)
      .options("/api/slides")
      .set("Origin", "https://untrusted.example")
      .set("Access-Control-Request-Method", "GET")
      .expect(403);
    expect(deniedResponse.body.error.code).toBe("CORS_NOT_ALLOWED");
  });
});
