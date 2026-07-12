import {
  CatalogSchema,
  SlideLibraryItemSchema,
  type LibraryValidationIssue,
  type LibraryValidationReport,
  type SlideLibraryItem
} from "@slide-library/shared";
import { readFile, realpath, stat } from "node:fs/promises";
import path from "node:path";

import { UnsafeAssetPathError } from "./storage/errors.js";
import { isPathInside, resolveContainedPath } from "./storage/pathSafety.js";

function zodIssuePath(pathParts: PropertyKey[]): string {
  return pathParts.reduce<string>((result, part) => {
    if (typeof part === "number") {
      return `${result}[${part}]`;
    }
    return result ? `${result}.${String(part)}` : String(part);
  }, "catalog.json");
}

function itemIdAt(catalog: unknown, pathParts: PropertyKey[]): string | undefined {
  const index = pathParts[0];
  if (!Array.isArray(catalog) || typeof index !== "number") {
    return undefined;
  }
  const candidate = catalog[index];
  if (typeof candidate !== "object" || candidate === null || !("id" in candidate)) {
    return undefined;
  }
  return typeof candidate.id === "string" ? candidate.id : undefined;
}

async function validateRegisteredFile(
  libraryRoot: string,
  realLibraryRoot: string,
  item: SlideLibraryItem,
  kind: "source" | "preview"
): Promise<LibraryValidationIssue | undefined> {
  const relativePath = kind === "source" ? item.sourceFile : item.previewFile;
  try {
    const candidate = resolveContainedPath(libraryRoot, relativePath);
    const [realCandidate, candidateStat] = await Promise.all([
      realpath(candidate),
      stat(candidate)
    ]);
    if (!candidateStat.isFile()) {
      return {
        itemId: item.id,
        path: relativePath,
        message: `Registered ${kind} asset is not a file`
      };
    }
    if (!isPathInside(realLibraryRoot, realCandidate)) {
      return {
        itemId: item.id,
        path: relativePath,
        message: `Registered ${kind} asset resolves outside the library root`
      };
    }
  } catch (error) {
    return {
      itemId: item.id,
      path: relativePath,
      message:
        error instanceof UnsafeAssetPathError
          ? error.message
          : `Registered ${kind} asset does not exist or cannot be read`
    };
  }

  return undefined;
}

export async function validateLibrary(
  libraryPath: string
): Promise<LibraryValidationReport> {
  const rootPath = path.resolve(libraryPath);
  let catalogPath: string;
  try {
    catalogPath = resolveContainedPath(rootPath, "catalog.json");
  } catch (error) {
    return {
      valid: false,
      itemCount: 0,
      errors: [
        {
          path: "catalog.json",
          message: error instanceof Error ? error.message : "Catalog path is unsafe"
        }
      ],
      warnings: []
    };
  }

  let untrustedCatalog: unknown;
  try {
    untrustedCatalog = JSON.parse(await readFile(catalogPath, "utf8")) as unknown;
  } catch (error) {
    return {
      valid: false,
      itemCount: 0,
      errors: [
        {
          path: "catalog.json",
          message:
            error instanceof SyntaxError
              ? "Catalog contains invalid JSON"
              : "Catalog file does not exist or cannot be read"
        }
      ],
      warnings: []
    };
  }

  const itemCount = Array.isArray(untrustedCatalog) ? untrustedCatalog.length : 0;
  const errors: LibraryValidationIssue[] = [];
  const warnings: LibraryValidationIssue[] = [];
  const catalogResult = CatalogSchema.safeParse(untrustedCatalog);
  if (!catalogResult.success) {
    errors.push(
      ...catalogResult.error.issues.map((issue) => {
        const itemId = itemIdAt(untrustedCatalog, issue.path);
        return {
          ...(itemId === undefined ? {} : { itemId }),
          path: zodIssuePath(issue.path),
          message: issue.message
        };
      })
    );
  }

  const individuallyValidItems: SlideLibraryItem[] = Array.isArray(untrustedCatalog)
    ? untrustedCatalog.flatMap((candidate) => {
        const parsed = SlideLibraryItemSchema.safeParse(candidate);
        return parsed.success ? [parsed.data] : [];
      })
    : [];

  const seenIds = new Set<string>();
  for (const item of individuallyValidItems) {
    if (seenIds.has(item.id)) {
      errors.push({
        itemId: item.id,
        path: "catalog.json",
        message: `Duplicate slide id '${item.id}'`
      });
    }
    seenIds.add(item.id);
  }

  let realLibraryRoot: string;
  try {
    realLibraryRoot = await realpath(rootPath);
  } catch {
    errors.push({ path: "catalog.json", message: "Library root does not exist or cannot be read" });
    return { valid: false, itemCount, errors, warnings };
  }

  for (const item of individuallyValidItems) {
    const fileIssues = await Promise.all([
      validateRegisteredFile(rootPath, realLibraryRoot, item, "source"),
      validateRegisteredFile(rootPath, realLibraryRoot, item, "preview")
    ]);
    errors.push(...fileIssues.filter((issue): issue is LibraryValidationIssue => issue !== undefined));
  }

  return { valid: errors.length === 0, itemCount, errors, warnings };
}
