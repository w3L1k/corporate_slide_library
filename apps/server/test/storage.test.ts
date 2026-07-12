import type { SlideLibraryItem } from "@slide-library/shared";
import { mkdir, mkdtemp, rm, stat, symlink, utimes, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";

import { validateLibrary } from "../src/libraryValidation.js";
import type { Logger } from "../src/logger.js";
import { FileSystemSlideStorage } from "../src/storage/FileSystemSlideStorage.js";
import {
  CatalogValidationError,
  UnsafeAssetPathError
} from "../src/storage/errors.js";
import { resolveContainedPath } from "../src/storage/pathSafety.js";

const silentLogger: Logger = {
  info: () => undefined,
  warn: () => undefined,
  error: () => undefined
};

function makeItem(overrides: Partial<SlideLibraryItem> = {}): SlideLibraryItem {
  return {
    id: "example-slide",
    title: "Example Slide",
    category: "Strategy",
    tags: ["Example"],
    version: "1.0",
    status: "approved",
    updatedAt: "2026-07-01T10:00:00.000Z",
    sourceFile: "slides/example-slide.pptx",
    previewFile: "previews/example-slide.png",
    ...overrides
  };
}

describe("FileSystemSlideStorage", () => {
  const temporaryRoots: string[] = [];

  afterEach(async () => {
    await Promise.all(
      temporaryRoots.splice(0).map((root) => rm(root, { recursive: true, force: true }))
    );
  });

  async function createLibrary(items: unknown): Promise<string> {
    const root = await mkdtemp(path.join(os.tmpdir(), "slide-storage-"));
    temporaryRoots.push(root);
    await Promise.all([
      mkdir(path.join(root, "slides")),
      mkdir(path.join(root, "previews"))
    ]);
    await writeFile(path.join(root, "catalog.json"), JSON.stringify(items, null, 2), "utf8");
    return root;
  }

  it("rejects catalog metadata containing traversal paths", async () => {
    const unsafeItem = {
      ...makeItem(),
      sourceFile: "../outside.pptx"
    };
    const root = await createLibrary([unsafeItem]);
    const storage = new FileSystemSlideStorage(root, { logger: silentLogger });

    await expect(storage.refresh()).rejects.toBeInstanceOf(CatalogValidationError);
  });

  it("accepts an empty catalog as a valid library state", async () => {
    const root = await createLibrary([]);
    const storage = new FileSystemSlideStorage(root, { logger: silentLogger });

    await expect(storage.refresh()).resolves.toEqual([]);
    await expect(validateLibrary(root)).resolves.toMatchObject({
      valid: true,
      itemCount: 0,
      errors: []
    });
  });

  it("refreshes cached metadata after the catalog mtime changes", async () => {
    const root = await createLibrary([makeItem()]);
    const catalogPath = path.join(root, "catalog.json");
    const storage = new FileSystemSlideStorage(root, { logger: silentLogger });
    expect((await storage.getCatalog())[0]?.title).toBe("Example Slide");

    const updatedItem = makeItem({ title: "Updated Example Slide" });
    await writeFile(catalogPath, JSON.stringify([updatedItem], null, 2), "utf8");
    const future = new Date(Date.now() + 2_000);
    await utimes(catalogPath, future, future);

    expect((await storage.getCatalog())[0]?.title).toBe("Updated Example Slide");
  });

  it("force-refreshes metadata even when the catalog signature is unchanged", async () => {
    const root = await createLibrary([makeItem()]);
    const catalogPath = path.join(root, "catalog.json");
    const storage = new FileSystemSlideStorage(root, { logger: silentLogger });
    expect((await storage.getCatalog())[0]?.title).toBe("Example Slide");

    const originalStat = await stat(catalogPath);
    await writeFile(
      catalogPath,
      JSON.stringify([makeItem({ title: "Changed Slide" })], null, 2),
      "utf8"
    );
    await utimes(catalogPath, originalStat.atime, originalStat.mtime);

    expect((await storage.refresh())[0]?.title).toBe("Changed Slide");
  });

  it("rejects duplicate ids during catalog loading", async () => {
    const root = await createLibrary([makeItem(), makeItem()]);
    const storage = new FileSystemSlideStorage(root, { logger: silentLogger });

    await expect(storage.refresh()).rejects.toMatchObject({
      name: "CatalogValidationError",
      issues: ["Duplicate id: example-slide"]
    });
  });

  it("reports missing registered assets during full library validation", async () => {
    const root = await createLibrary([makeItem()]);
    await writeFile(path.join(root, "slides", "example-slide.pptx"), "pptx");

    const report = await validateLibrary(root);
    expect(report.valid).toBe(false);
    expect(report.itemCount).toBe(1);
    expect(report.errors).toContainEqual({
      itemId: "example-slide",
      path: "previews/example-slide.png",
      message: "Registered preview asset does not exist or cannot be read"
    });
  });

  it("blocks registered assets that resolve through a link outside the library", async () => {
    const root = await createLibrary([
      makeItem({ previewFile: "linked-previews/example-slide.png" })
    ]);
    const outsideRoot = await mkdtemp(path.join(os.tmpdir(), "slide-storage-outside-"));
    temporaryRoots.push(outsideRoot);
    await Promise.all([
      writeFile(path.join(root, "slides", "example-slide.pptx"), "pptx"),
      writeFile(path.join(outsideRoot, "example-slide.png"), "png")
    ]);
    await symlink(
      outsideRoot,
      path.join(root, "linked-previews"),
      process.platform === "win32" ? "junction" : "dir"
    );

    const storage = new FileSystemSlideStorage(root, { logger: silentLogger });
    await expect(storage.getPreview("example-slide")).rejects.toBeInstanceOf(
      UnsafeAssetPathError
    );

    const report = await validateLibrary(root);
    expect(report.errors).toContainEqual({
      itemId: "example-slide",
      path: "linked-previews/example-slide.png",
      message: "Registered preview asset resolves outside the library root"
    });
  });

  it.each([
    "../outside.pptx",
    "slides//example.pptx",
    "C:/outside.pptx",
    "slides/example\u0000.pptx"
  ])("rejects unsafe relative asset paths: %s", (registeredPath) => {
    expect(() => resolveContainedPath("C:/library", registeredPath)).toThrow(
      UnsafeAssetPathError
    );
  });
});
