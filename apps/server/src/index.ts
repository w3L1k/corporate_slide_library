import { createApp } from "./app.js";
import { loadConfig } from "./config.js";
import { logger } from "./logger.js";
import { FileSystemSlideStorage } from "./storage/FileSystemSlideStorage.js";

async function start(): Promise<void> {
  const config = loadConfig();
  logger.info("Starting slide library server", {
    host: config.host,
    port: config.port,
    libraryPath: config.libraryPath,
    adminReindexEnabled: config.enableAdminReindex
  });

  const storage = new FileSystemSlideStorage(config.libraryPath, { logger });
  const catalog = await storage.refresh();
  logger.info("Initial slide catalog is ready", { itemCount: catalog.length });

  const app = createApp({ storage, config, logger });
  const server = app.listen(config.port, config.host, () => {
    logger.info("Slide library server is listening", {
      host: config.host,
      port: config.port
    });
  });

  server.on("error", (error) => {
    logger.error("Slide library server failed", { error });
    process.exitCode = 1;
  });

  const shutdown = (signal: string): void => {
    logger.info("Stopping slide library server", { signal });
    server.close((error) => {
      if (error) {
        logger.error("Slide library server shutdown failed", { error });
        process.exitCode = 1;
      }
    });
  };

  process.once("SIGINT", () => shutdown("SIGINT"));
  process.once("SIGTERM", () => shutdown("SIGTERM"));
}

start().catch((error: unknown) => {
  logger.error("Slide library server startup failed", { error });
  process.exitCode = 1;
});
