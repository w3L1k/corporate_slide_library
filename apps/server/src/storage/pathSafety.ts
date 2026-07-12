import path from "node:path";

import { UnsafeAssetPathError } from "./errors.js";

function containsControlCharacter(value: string): boolean {
  return [...value].some((character) => {
    const codePoint = character.codePointAt(0);
    return codePoint !== undefined && (codePoint < 32 || codePoint === 127);
  });
}

export function isPathInside(rootPath: string, candidatePath: string): boolean {
  const relative = path.relative(rootPath, candidatePath);
  return (
    relative === "" ||
    (!relative.startsWith(`..${path.sep}`) &&
      relative !== ".." &&
      !path.isAbsolute(relative))
  );
}

export function resolveContainedPath(
  rootPath: string,
  relativePath: string
): string {
  if (!relativePath || path.isAbsolute(relativePath)) {
    throw new UnsafeAssetPathError("Asset path must be relative to the library root");
  }

  const normalizedSegments = relativePath.split("/");
  if (
    relativePath.includes("\\") ||
    containsControlCharacter(relativePath) ||
    normalizedSegments.some(
      (segment) =>
        segment === "" || segment === "." || segment === ".." || segment.includes(":")
    )
  ) {
    throw new UnsafeAssetPathError("Asset path contains an unsafe path segment");
  }

  const resolvedRoot = path.resolve(rootPath);
  const candidate = path.resolve(resolvedRoot, ...normalizedSegments);
  if (!isPathInside(resolvedRoot, candidate) || candidate === resolvedRoot) {
    throw new UnsafeAssetPathError("Asset path resolves outside the library root");
  }

  return candidate;
}
