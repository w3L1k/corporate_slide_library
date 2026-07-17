import {
  PersonalAssetCatalogSchema,
  PersonalAssetSchema,
  type PersonalAsset,
  type PersonalAssetKind,
  type PersonalAssetRecord
} from "@slide-library/shared";
import { randomUUID } from "node:crypto";
import {
  mkdir,
  readFile,
  realpath,
  rename,
  rm,
  stat,
  writeFile
} from "node:fs/promises";
import path from "node:path";

import { isPathInside, resolveContainedPath } from "./pathSafety.js";

export interface AddPersonalAssetInput {
  title: string;
  kind: PersonalAssetKind;
  originalFileName: string;
  mimeType: string;
  extension: string;
  data: Buffer;
}

export interface PersonalAssetFile {
  item: PersonalAsset;
  data: Buffer;
}

const hasErrorCode = (error: unknown, code: string): boolean =>
  typeof error === "object" &&
  error !== null &&
  "code" in error &&
  error.code === code;

const toPublicItem = (record: PersonalAssetRecord): PersonalAsset =>
  PersonalAssetSchema.parse({
    id: record.id,
    title: record.title,
    kind: record.kind,
    fileName: record.fileName,
    mimeType: record.mimeType,
    size: record.size,
    createdAt: record.createdAt
  });

export class PersonalAssetStorage {
  readonly rootPath: string;
  readonly filesPath: string;
  readonly catalogPath: string;

  private writeQueue: Promise<void> = Promise.resolve();

  constructor(libraryPath: string) {
    this.rootPath = resolveContainedPath(path.resolve(libraryPath), "personal-assets");
    this.filesPath = resolveContainedPath(this.rootPath, "files");
    this.catalogPath = resolveContainedPath(this.rootPath, "index.json");
  }

  async initialize(): Promise<void> {
    await mkdir(this.filesPath, { recursive: true });
    try {
      await stat(this.catalogPath);
    } catch {
      await writeFile(this.catalogPath, "[]\n", { encoding: "utf8", flag: "wx" }).catch(
        (error: unknown) => {
          if (
            typeof error !== "object" ||
            error === null ||
            !("code" in error) ||
            error.code !== "EEXIST"
          ) {
            throw error;
          }
        }
      );
    }
  }

  async list(): Promise<PersonalAsset[]> {
    const records = await this.readCatalog();
    return records
      .map(toPublicItem)
      .sort((left, right) => Date.parse(right.createdAt) - Date.parse(left.createdAt));
  }

  async add(input: AddPersonalAssetInput): Promise<PersonalAsset> {
    await this.initialize();

    const id = randomUUID();
    const storedFile = `files/${id}${input.extension}`;
    const record = PersonalAssetCatalogSchema.element.parse({
      id,
      title: input.title,
      kind: input.kind,
      fileName: input.originalFileName,
      mimeType: input.mimeType,
      size: input.data.byteLength,
      createdAt: new Date().toISOString(),
      storedFile
    });
    const targetPath = resolveContainedPath(this.rootPath, storedFile);

    await writeFile(targetPath, input.data, { flag: "wx" });
    try {
      await this.enqueueWrite(async () => {
        const records = await this.readCatalog();
        records.push(record);
        await this.writeCatalog(records);
      });
    } catch (error) {
      await rm(targetPath, { force: true });
      throw error;
    }

    return toPublicItem(record);
  }

  async getFile(id: string): Promise<PersonalAssetFile | undefined> {
    const records = await this.readCatalog();
    const record = records.find((item) => item.id === id);
    if (!record) {
      return undefined;
    }

    const assetPath = resolveContainedPath(this.rootPath, record.storedFile);
    const [realRoot, realAsset, assetStat] = await Promise.all([
      realpath(this.rootPath),
      realpath(assetPath),
      stat(assetPath)
    ]);
    if (!assetStat.isFile() || !isPathInside(realRoot, realAsset)) {
      throw new Error("Registered personal asset is outside the personal library");
    }

    return { item: toPublicItem(record), data: await readFile(realAsset) };
  }

  async remove(id: string): Promise<PersonalAsset | undefined> {
    let removedItem: PersonalAsset | undefined;

    await this.enqueueWrite(async () => {
      const records = await this.readCatalog();
      const recordIndex = records.findIndex((item) => item.id === id);
      if (recordIndex < 0) {
        return;
      }

      const record = records[recordIndex]!;
      const assetPath = resolveContainedPath(this.rootPath, record.storedFile);
      const temporaryPath = resolveContainedPath(
        this.filesPath,
        `.delete-${randomUUID()}.tmp`
      );
      let movedFile = false;

      try {
        const [realRoot, realAsset, assetStat] = await Promise.all([
          realpath(this.rootPath),
          realpath(assetPath),
          stat(assetPath)
        ]);
        if (!assetStat.isFile() || !isPathInside(realRoot, realAsset)) {
          throw new Error("Registered personal asset is outside the personal library");
        }

        await rename(assetPath, temporaryPath);
        movedFile = true;
      } catch (error) {
        if (!hasErrorCode(error, "ENOENT")) {
          throw error;
        }
      }

      const nextRecords = records.filter((_, index) => index !== recordIndex);
      try {
        await this.writeCatalog(nextRecords);
      } catch (error) {
        if (movedFile) {
          await rename(temporaryPath, assetPath);
        }
        throw error;
      }

      if (movedFile) {
        await rm(temporaryPath, { force: true });
      }
      removedItem = toPublicItem(record);
    });

    return removedItem;
  }

  private async readCatalog(): Promise<PersonalAssetRecord[]> {
    await this.initialize();
    const rawCatalog = await readFile(this.catalogPath, "utf8");
    return PersonalAssetCatalogSchema.parse(JSON.parse(rawCatalog) as unknown);
  }

  private async writeCatalog(records: PersonalAssetRecord[]): Promise<void> {
    const validated = PersonalAssetCatalogSchema.parse(records);
    const temporaryPath = resolveContainedPath(
      this.rootPath,
      `index.${randomUUID()}.tmp`
    );
    await writeFile(temporaryPath, `${JSON.stringify(validated, null, 2)}\n`, "utf8");
    await rename(temporaryPath, this.catalogPath);
  }

  private async enqueueWrite(operation: () => Promise<void>): Promise<void> {
    const previous = this.writeQueue;
    let release: (() => void) | undefined;
    this.writeQueue = new Promise<void>((resolve) => {
      release = resolve;
    });

    await previous;
    try {
      await operation();
    } finally {
      release?.();
    }
  }
}
