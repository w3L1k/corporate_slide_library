import { CatalogSchema, type SlideLibraryItem } from "@slide-library/shared";
import { readFile, realpath, stat } from "node:fs/promises";
import path from "node:path";

import type { Logger } from "../logger.js";
import { logger as defaultLogger } from "../logger.js";
import type { SlideStorage } from "./SlideStorage.js";
import {
  AssetNotFoundError,
  CatalogReadError,
  CatalogValidationError,
  SlideNotFoundError,
  UnsafeAssetPathError,
  type AssetKind
} from "./errors.js";
import { isPathInside, resolveContainedPath } from "./pathSafety.js";

interface FileSystemSlideStorageOptions {
  catalogFileName?: string;
  logger?: Logger;
}

interface CatalogSignature {
  mtimeMs: number;
  size: number;
}

const PREVIEW_EXTENSION_PATTERN = /\.(?:png|jpe?g|webp)$/i;

function signaturesEqual(
  left: CatalogSignature | undefined,
  right: CatalogSignature
): boolean {
  return left?.mtimeMs === right.mtimeMs && left.size === right.size;
}

function formatZodPath(pathParts: PropertyKey[]): string {
  if (pathParts.length === 0) {
    return "catalog";
  }

  return pathParts.reduce<string>((result, part) => {
    if (typeof part === "number") {
      return `${result}[${part}]`;
    }
    return result ? `${result}.${String(part)}` : String(part);
  }, "catalog");
}

export class FileSystemSlideStorage implements SlideStorage {
  readonly rootPath: string;
  readonly catalogPath: string;

  private readonly logger: Logger;
  private catalog: SlideLibraryItem[] = [];
  private itemsById = new Map<string, SlideLibraryItem>();
  private catalogSignature: CatalogSignature | undefined;
  private loadPromise: Promise<void> | undefined;

  constructor(rootPath: string, options: FileSystemSlideStorageOptions = {}) {
    this.rootPath = path.resolve(rootPath);
    this.logger = options.logger ?? defaultLogger;
    this.catalogPath = resolveContainedPath(
      this.rootPath,
      options.catalogFileName ?? "catalog.json"
    );

    if (path.extname(this.catalogPath).toLowerCase() !== ".json") {
      throw new UnsafeAssetPathError("Catalog must be a JSON file inside the library root");
    }
  }

  async getCatalog(): Promise<SlideLibraryItem[]> {
    await this.ensureFresh();
    return [...this.catalog];
  }

  async getItem(id: string): Promise<SlideLibraryItem | undefined> {
    await this.ensureFresh();
    return this.itemsById.get(id);
  }

  async getSlide(id: string): Promise<Buffer> {
    return this.readRegisteredAsset(id, "slide");
  }

  async getPreview(id: string): Promise<Buffer> {
    return this.readRegisteredAsset(id, "preview");
  }

  async refresh(): Promise<SlideLibraryItem[]> {
    await this.loadCatalog(true);
    return [...this.catalog];
  }

  private async ensureFresh(): Promise<void> {
    let signature: CatalogSignature;
    try {
      signature = await this.getCatalogSignature();
    } catch (error) {
      throw new CatalogReadError("The slide catalog could not be read", { cause: error });
    }

    if (!signaturesEqual(this.catalogSignature, signature)) {
      await this.loadCatalog(false, signature);
    }
  }

  private async loadCatalog(
    force: boolean,
    knownSignature?: CatalogSignature
  ): Promise<void> {
    if (this.loadPromise) {
      await this.loadPromise;
      if (force) {
        await this.loadCatalog(false);
      }
      return;
    }

    this.loadPromise = this.performCatalogLoad(force, knownSignature);
    try {
      await this.loadPromise;
    } finally {
      this.loadPromise = undefined;
    }
  }

  private async performCatalogLoad(
    force: boolean,
    knownSignature?: CatalogSignature
  ): Promise<void> {
    let signature = knownSignature;
    let rawCatalog: string;

    try {
      signature ??= await this.getCatalogSignature();
      if (!force && signaturesEqual(this.catalogSignature, signature)) {
        return;
      }
      rawCatalog = await readFile(this.catalogPath, "utf8");
    } catch (error) {
      this.logger.error("Catalog read failed", { error });
      throw new CatalogReadError("The slide catalog could not be read", { cause: error });
    }

    let untrustedCatalog: unknown;
    try {
      untrustedCatalog = JSON.parse(rawCatalog) as unknown;
    } catch (error) {
      this.logger.error("Catalog JSON parsing failed", { error });
      throw new CatalogValidationError("The slide catalog contains invalid JSON");
    }

    const result = CatalogSchema.safeParse(untrustedCatalog);
    if (!result.success) {
      const issues = result.error.issues.map(
        (issue) => `${formatZodPath(issue.path)}: ${issue.message}`
      );
      this.logger.error("Catalog schema validation failed", { issues });
      throw new CatalogValidationError(
        "The slide catalog does not match the required schema",
        issues
      );
    }

    const duplicateIds = result.data
      .map((item) => item.id)
      .filter((id, index, ids) => ids.indexOf(id) !== index);
    if (duplicateIds.length > 0) {
      const uniqueDuplicates = [...new Set(duplicateIds)];
      this.logger.error("Catalog contains duplicate slide ids", {
        duplicateIds: uniqueDuplicates
      });
      throw new CatalogValidationError(
        "The slide catalog contains duplicate ids",
        uniqueDuplicates.map((id) => `Duplicate id: ${id}`)
      );
    }

    // Replace the cache only after the entire new catalog has passed validation.
    this.catalog = result.data;
    this.itemsById = new Map(result.data.map((item) => [item.id, item]));
    // Keep the signature observed before the read. If the file changed during the
    // read, the next request observes a mismatch and safely reloads it again.
    this.catalogSignature = signature;
    this.logger.info("Catalog loaded", { itemCount: this.catalog.length });
  }

  private async getCatalogSignature(): Promise<CatalogSignature> {
    const fileStat = await stat(this.catalogPath);
    if (!fileStat.isFile()) {
      throw new Error("Catalog path is not a file");
    }
    return { mtimeMs: fileStat.mtimeMs, size: fileStat.size };
  }

  private async readRegisteredAsset(id: string, kind: AssetKind): Promise<Buffer> {
    await this.ensureFresh();
    const item = this.itemsById.get(id);
    if (!item) {
      throw new SlideNotFoundError(id);
    }

    const registeredPath = kind === "slide" ? item.sourceFile : item.previewFile;
    if (
      (kind === "slide" && path.extname(registeredPath).toLowerCase() !== ".pptx") ||
      (kind === "preview" && !PREVIEW_EXTENSION_PATTERN.test(registeredPath))
    ) {
      throw new UnsafeAssetPathError(`Registered ${kind} asset has an unsupported extension`);
    }

    try {
      const assetPath = resolveContainedPath(this.rootPath, registeredPath);
      const [realRoot, realAsset, assetStat] = await Promise.all([
        realpath(this.rootPath),
        realpath(assetPath),
        stat(assetPath)
      ]);
      if (!assetStat.isFile() || !isPathInside(realRoot, realAsset)) {
        throw new UnsafeAssetPathError(`Registered ${kind} asset is outside the library root`);
      }
      return await readFile(realAsset);
    } catch (error) {
      if (error instanceof UnsafeAssetPathError) {
        throw error;
      }
      throw new AssetNotFoundError(id, kind, { cause: error });
    }
  }
}
