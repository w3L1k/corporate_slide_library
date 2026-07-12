import path from "node:path";
import { pathToFileURL } from "node:url";

import { DEFAULT_DEMO_LIBRARY_PATH } from "../apps/server/src/config.js";

export interface CliIo {
  writeOutput(message: string): void;
  writeError(message: string): void;
}

export const consoleCliIo: CliIo = {
  writeOutput: (message) => console.log(message),
  writeError: (message) => console.error(message)
};

export function isDirectCliExecution(
  moduleUrl: string,
  argumentsList: string[] = process.argv
): boolean {
  const entryPath = argumentsList[1];
  return entryPath !== undefined && moduleUrl === pathToFileURL(entryPath).href;
}

export function getLibraryPath(
  argumentsList: string[] = process.argv.slice(2),
  environment: NodeJS.ProcessEnv = process.env
): string {
  let configuredPath: string | undefined;

  const setConfiguredPath = (rawPath: string): void => {
    const candidate = rawPath.trim();
    if (!candidate) {
      throw new Error("--path requires a non-empty directory argument");
    }
    if (configuredPath !== undefined) {
      throw new Error("Library path may only be provided once");
    }
    configuredPath = candidate;
  };

  for (let index = 0; index < argumentsList.length; index += 1) {
    const argument = argumentsList[index];
    if (argument === "--path") {
      const nextArgument = argumentsList[index + 1];
      if (!nextArgument || nextArgument.startsWith("--")) {
        throw new Error("--path requires a directory argument");
      }
      setConfiguredPath(nextArgument);
      index += 1;
      continue;
    }
    if (argument?.startsWith("--path=")) {
      setConfiguredPath(argument.slice("--path=".length));
      continue;
    }
    if (argument?.startsWith("--")) {
      throw new Error(`Unknown option: ${argument}`);
    }
    if (argument) {
      setConfiguredPath(argument);
      continue;
    }
  }

  return path.resolve(
    configuredPath ||
      environment.SLIDE_LIBRARY_PATH?.trim() ||
      DEFAULT_DEMO_LIBRARY_PATH
  );
}
