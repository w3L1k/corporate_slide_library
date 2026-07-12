import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";

import { DEFAULT_DEMO_LIBRARY_PATH } from "../src/config.js";
import {
  getLibraryPath,
  type CliIo
} from "../../../tools/library-cli.js";
import { runReindexLibrary } from "../../../tools/reindex-library.js";
import { runValidateLibrary } from "../../../tools/validate-library.js";

const temporaryRoots: string[] = [];

async function createLibrary(includePreview = true): Promise<string> {
  const root = await mkdtemp(path.join(os.tmpdir(), "slide-cli-"));
  temporaryRoots.push(root);
  await Promise.all([
    mkdir(path.join(root, "slides")),
    mkdir(path.join(root, "previews"))
  ]);
  await writeFile(
    path.join(root, "catalog.json"),
    JSON.stringify([
      {
        id: "example-slide",
        title: "Example Slide",
        category: "Strategy",
        tags: ["Example"],
        version: "1.0",
        status: "approved",
        updatedAt: "2026-07-01T10:00:00.000Z",
        sourceFile: "slides/example-slide.pptx",
        previewFile: "previews/example-slide.png"
      }
    ]),
    "utf8"
  );
  await writeFile(path.join(root, "slides", "example-slide.pptx"), "pptx");
  if (includePreview) {
    await writeFile(path.join(root, "previews", "example-slide.png"), "png");
  }
  return root;
}

function captureIo(): {
  io: CliIo;
  output: string[];
  errors: string[];
} {
  const output: string[] = [];
  const errors: string[] = [];
  return {
    io: {
      writeOutput: (message) => output.push(message),
      writeError: (message) => errors.push(message)
    },
    output,
    errors
  };
}

afterEach(async () => {
  await Promise.all(
    temporaryRoots.splice(0).map((root) => rm(root, { recursive: true, force: true }))
  );
});

describe("library CLI path parsing", () => {
  it("uses argument, environment, and demo paths in precedence order", () => {
    expect(
      getLibraryPath(["--path", "./argument-library"], {
        SLIDE_LIBRARY_PATH: "./environment-library"
      })
    ).toBe(path.resolve("./argument-library"));
    expect(getLibraryPath([], { SLIDE_LIBRARY_PATH: "./environment-library" })).toBe(
      path.resolve("./environment-library")
    );
    expect(getLibraryPath([], {})).toBe(DEFAULT_DEMO_LIBRARY_PATH);
  });

  it("supports --path=value and a single positional path", () => {
    expect(getLibraryPath(["--path=./library"], {})).toBe(path.resolve("./library"));
    expect(getLibraryPath(["./library"], {})).toBe(path.resolve("./library"));
  });

  it.each([
    [["--path"], /requires a directory/],
    [["--path="], /non-empty directory/],
    [["--unknown"], /Unknown option/],
    [["first", "second"], /only be provided once/],
    [["--path", "first", "--path", "second"], /only be provided once/]
  ])("rejects ambiguous or invalid arguments: %o", (argumentsList, expectedMessage) => {
    expect(() => getLibraryPath(argumentsList, {})).toThrow(expectedMessage);
  });

  it("validates and reindexes a complete registered library", async () => {
    const root = await createLibrary();
    const validationCapture = captureIo();
    const reindexCapture = captureIo();

    await expect(
      runValidateLibrary(["--path", root], {}, validationCapture.io)
    ).resolves.toBe(0);
    expect(validationCapture.errors).toEqual([]);
    expect(JSON.parse(validationCapture.output[0] ?? "{}")).toMatchObject({
      libraryPath: root,
      valid: true,
      itemCount: 1
    });

    await expect(
      runReindexLibrary(["--path", root], {}, reindexCapture.io)
    ).resolves.toBe(0);
    expect(reindexCapture.errors).toEqual([]);
    expect(JSON.parse(reindexCapture.output[0] ?? "{}")).toMatchObject({
      status: "ok",
      libraryPath: root,
      itemCount: 1,
      runtimeReload: "automatic-on-next-api-read-or-gated-admin-reindex"
    });
  });

  it("returns a non-zero result and structured issues for an invalid library", async () => {
    const root = await createLibrary(false);
    const validationCapture = captureIo();
    const reindexCapture = captureIo();

    await expect(
      runValidateLibrary([root], {}, validationCapture.io)
    ).resolves.toBe(1);
    expect(JSON.parse(validationCapture.output[0] ?? "{}")).toMatchObject({
      valid: false,
      itemCount: 1
    });

    await expect(runReindexLibrary([root], {}, reindexCapture.io)).resolves.toBe(1);
    expect(JSON.parse(reindexCapture.errors[0] ?? "{}")).toMatchObject({
      valid: false,
      itemCount: 1
    });
  });
});
