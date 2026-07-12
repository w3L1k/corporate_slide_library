import { validateLibrary } from "../apps/server/src/libraryValidation.js";
import type { Logger } from "../apps/server/src/logger.js";
import { FileSystemSlideStorage } from "../apps/server/src/storage/FileSystemSlideStorage.js";
import {
  consoleCliIo,
  getLibraryPath,
  isDirectCliExecution,
  type CliIo
} from "./library-cli.js";

const cliLogger: Logger = {
  info: () => undefined,
  warn: () => undefined,
  error: () => undefined
};

export async function runReindexLibrary(
  argumentsList: string[] = process.argv.slice(2),
  environment: NodeJS.ProcessEnv = process.env,
  io: CliIo = consoleCliIo
): Promise<number> {
  try {
    const libraryPath = getLibraryPath(argumentsList, environment);
    const validation = await validateLibrary(libraryPath);

    if (!validation.valid) {
      io.writeError(JSON.stringify({ libraryPath, ...validation }, null, 2));
      return 1;
    }

    // This is a read-only rebuild and loader smoke test. A running server notices
    // catalog mtime changes automatically; its gated admin route forces an immediate
    // reload when a content owner needs one without waiting for the next API read.
    const storage = new FileSystemSlideStorage(libraryPath, { logger: cliLogger });
    const items = await storage.refresh();
    io.writeOutput(
      JSON.stringify(
        {
          status: "ok",
          libraryPath,
          itemCount: items.length,
          refreshedAt: new Date().toISOString(),
          runtimeReload: "automatic-on-next-api-read-or-gated-admin-reindex"
        },
        null,
        2
      )
    );
    return 0;
  } catch (error) {
    io.writeError(
      JSON.stringify({
        status: "error",
        error: error instanceof Error ? error.message : "Library reindex failed"
      })
    );
    return 1;
  }
}

if (isDirectCliExecution(import.meta.url)) {
  void runReindexLibrary().then((exitCode) => {
    process.exitCode = exitCode;
  });
}
