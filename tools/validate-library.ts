import { validateLibrary } from "../apps/server/src/libraryValidation.js";
import {
  consoleCliIo,
  getLibraryPath,
  isDirectCliExecution,
  type CliIo
} from "./library-cli.js";

export async function runValidateLibrary(
  argumentsList: string[] = process.argv.slice(2),
  environment: NodeJS.ProcessEnv = process.env,
  io: CliIo = consoleCliIo
): Promise<number> {
  try {
    const libraryPath = getLibraryPath(argumentsList, environment);
    const report = await validateLibrary(libraryPath);

    io.writeOutput(
      JSON.stringify(
        {
          libraryPath,
          ...report
        },
        null,
        2
      )
    );

    return report.valid ? 0 : 1;
  } catch (error) {
    io.writeError(
      JSON.stringify({
        valid: false,
        error: error instanceof Error ? error.message : "Library validation failed"
      })
    );
    return 1;
  }
}

if (isDirectCliExecution(import.meta.url)) {
  void runValidateLibrary().then((exitCode) => {
    process.exitCode = exitCode;
  });
}
