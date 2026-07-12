import {
  slideStatusValues,
  type ApiErrorResponse,
  type HealthResponse,
  type ReindexResponse,
  type SlideLibraryItem,
  type SlideListResponse,
  type SlideStatus
} from "@slide-library/shared";
import cors from "cors";
import { randomUUID } from "node:crypto";
import path from "node:path";
import express, {
  type ErrorRequestHandler,
  type Request,
  type RequestHandler,
  type Response
} from "express";

import type { ServerConfig } from "./config.js";
import { ApiError } from "./errors.js";
import type { Logger } from "./logger.js";
import { logger as defaultLogger } from "./logger.js";
import type { SlideStorage } from "./storage/SlideStorage.js";
import {
  AssetNotFoundError,
  CatalogReadError,
  CatalogValidationError,
  SlideNotFoundError,
  UnsafeAssetPathError
} from "./storage/errors.js";

export interface CreateAppOptions {
  storage: SlideStorage;
  config: Pick<ServerConfig, "corsOrigins" | "enableAdminReindex">;
  logger?: Logger;
}

const SLIDE_ID_PATTERN = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;
const POWERPOINT_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";

function asyncRoute(handler: RequestHandler): RequestHandler {
  return (request, response, next) => {
    Promise.resolve(handler(request, response, next)).catch(next);
  };
}

function getQueryString(value: unknown, name: string): string | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== "string") {
    throw new ApiError(400, "INVALID_QUERY", `${name} must be provided once as text`);
  }

  const trimmed = value.trim();
  return trimmed || undefined;
}

function normalizeForSearch(value: string): string {
  return value.normalize("NFKC").toLowerCase();
}

function matchesSearch(item: SlideLibraryItem, query: string): boolean {
  const searchableValues = [
    item.title,
    item.description,
    item.category,
    ...item.tags,
    item.searchText
  ];

  return searchableValues.some(
    (value) => value !== undefined && normalizeForSearch(value).includes(query)
  );
}

function requireValidId(request: Request): string {
  const id = request.params.id;
  if (typeof id !== "string" || !SLIDE_ID_PATTERN.test(id)) {
    throw new ApiError(400, "INVALID_SLIDE_ID", "Slide id has an invalid format");
  }
  return id;
}

function previewContentType(fileName: string): string {
  switch (path.extname(fileName).toLowerCase()) {
    case ".png":
      return "image/png";
    case ".jpg":
    case ".jpeg":
      return "image/jpeg";
    case ".webp":
      return "image/webp";
    default:
      return "application/octet-stream";
  }
}

function sendApiError(
  response: Response<ApiErrorResponse>,
  status: number,
  code: string,
  message: string
): void {
  response.status(status).json({ error: { code, message } });
}

function getHttpErrorStatus(error: unknown): number | undefined {
  if (typeof error !== "object" || error === null || !("status" in error)) {
    return undefined;
  }

  const status = error.status;
  return typeof status === "number" && Number.isInteger(status) ? status : undefined;
}

export function createApp(options: CreateAppOptions): express.Express {
  const { storage, config } = options;
  const log = options.logger ?? defaultLogger;
  const app = express();

  app.disable("x-powered-by");
  app.set("query parser", "simple");
  app.use((request, response, next) => {
    const startedAt = performance.now();
    const requestId = randomUUID();
    response.setHeader("X-Request-Id", requestId);
    response.setHeader("X-Content-Type-Options", "nosniff");
    response.once("finish", () => {
      log.info("API request completed", {
        requestId,
        method: request.method,
        path: request.path,
        status: response.statusCode,
        durationMs: Math.round((performance.now() - startedAt) * 10) / 10
      });
    });
    next();
  });
  app.use(
    cors({
      origin(origin, callback) {
        if (
          origin === undefined ||
          config.corsOrigins.includes("*") ||
          config.corsOrigins.includes(origin)
        ) {
          callback(null, true);
          return;
        }
        callback(new ApiError(403, "CORS_NOT_ALLOWED", "This origin is not allowed"));
      },
      methods: ["GET", "POST", "OPTIONS"],
      optionsSuccessStatus: 204
    })
  );
  app.use(express.json({ limit: "16kb" }));

  app.get("/api/health", (_request, response: Response<HealthResponse>) => {
    response.json({ status: "ok" });
  });

  app.get(
    "/api/slides",
    asyncRoute(async (request, response: Response<SlideListResponse>) => {
      const q = getQueryString(request.query.q, "q");
      const category = getQueryString(request.query.category, "category");
      const statusValue = getQueryString(request.query.status, "status");
      const normalizedStatus = statusValue?.toLowerCase();

      if (
        normalizedStatus !== undefined &&
        !slideStatusValues.includes(normalizedStatus as SlideStatus)
      ) {
        throw new ApiError(
          400,
          "INVALID_STATUS",
          `status must be one of: ${slideStatusValues.join(", ")}`
        );
      }

      const catalog = await storage.getCatalog();
      const availableCategories = [...new Set(catalog.map((item) => item.category))].sort(
        (left, right) => left.localeCompare(right, undefined, { sensitivity: "base" })
      );
      const normalizedQuery = q === undefined ? undefined : normalizeForSearch(q);
      const normalizedCategory =
        category === undefined ? undefined : normalizeForSearch(category);
      const items = catalog.filter((item) => {
        if (
          normalizedQuery !== undefined &&
          !matchesSearch(item, normalizedQuery)
        ) {
          return false;
        }
        if (
          normalizedCategory !== undefined &&
          normalizeForSearch(item.category) !== normalizedCategory
        ) {
          return false;
        }
        return normalizedStatus === undefined || item.status === normalizedStatus;
      });

      response.json({ items, total: items.length, availableCategories });
    })
  );

  app.get(
    "/api/slides/:id",
    asyncRoute(async (request, response: Response<SlideLibraryItem | ApiErrorResponse>) => {
      const id = requireValidId(request);
      const item = await storage.getItem(id);
      if (!item) {
        throw new SlideNotFoundError(id);
      }
      response.json(item);
    })
  );

  app.get(
    "/api/slides/:id/preview",
    asyncRoute(async (request, response) => {
      const id = requireValidId(request);
      const item = await storage.getItem(id);
      if (!item) {
        throw new SlideNotFoundError(id);
      }
      const preview = await storage.getPreview(id);
      response.setHeader("Content-Type", previewContentType(item.previewFile));
      response.setHeader("Cache-Control", "public, max-age=0, must-revalidate");
      response.send(preview);
    })
  );

  app.get(
    "/api/slides/:id/file",
    asyncRoute(async (request, response) => {
      const id = requireValidId(request);
      const item = await storage.getItem(id);
      if (!item) {
        throw new SlideNotFoundError(id);
      }
      const slide = await storage.getSlide(id);
      response.setHeader("Content-Type", POWERPOINT_CONTENT_TYPE);
      response.setHeader("Content-Disposition", `attachment; filename="${item.id}.pptx"`);
      response.setHeader("Cache-Control", "private, max-age=0, must-revalidate");
      response.send(slide);
    })
  );

  if (config.enableAdminReindex) {
    app.post(
      "/api/admin/reindex",
      asyncRoute(async (_request, response: Response<ReindexResponse>) => {
        const catalog = await storage.refresh();
        response.json({
          status: "ok",
          itemCount: catalog.length,
          refreshedAt: new Date().toISOString()
        });
      })
    );
  }

  app.use("/api", (_request, response: Response<ApiErrorResponse>) => {
    sendApiError(response, 404, "API_ROUTE_NOT_FOUND", "The requested API route does not exist");
  });

  const errorHandler: ErrorRequestHandler = (error, request, response, _next) => {
    if (response.headersSent) {
      _next(error);
      return;
    }

    const requestId = response.getHeader("X-Request-Id");

    if (error instanceof ApiError) {
      if (error.status >= 500) {
        log.error("API request failed", { requestId, path: request.path, error });
      }
      sendApiError(response, error.status, error.code, error.message);
      return;
    }
    if (error instanceof SlideNotFoundError) {
      sendApiError(response, 404, "SLIDE_NOT_FOUND", "The requested slide was not found");
      return;
    }
    if (error instanceof AssetNotFoundError) {
      const isPreview = error.kind === "preview";
      log.warn("Registered slide asset is unavailable", {
        requestId,
        path: request.path,
        slideId: error.id,
        assetKind: error.kind
      });
      sendApiError(
        response,
        404,
        isPreview ? "PREVIEW_NOT_FOUND" : "SLIDE_FILE_NOT_FOUND",
        isPreview
          ? "The preview for this slide is unavailable"
          : "The PowerPoint file for this slide is unavailable"
      );
      return;
    }
    if (error instanceof CatalogValidationError) {
      log.error("API request failed because the catalog is invalid", {
        requestId,
        path: request.path,
        error
      });
      sendApiError(
        response,
        500,
        "CATALOG_INVALID",
        "The slide library catalog is invalid. Ask the content owner to validate it."
      );
      return;
    }
    if (error instanceof CatalogReadError) {
      log.error("API request failed because the catalog is unavailable", {
        requestId,
        path: request.path,
        error
      });
      sendApiError(
        response,
        503,
        "CATALOG_UNAVAILABLE",
        "The slide library catalog is temporarily unavailable"
      );
      return;
    }
    if (error instanceof UnsafeAssetPathError) {
      log.error("Blocked an unsafe registered asset path", {
        requestId,
        path: request.path,
        error
      });
      sendApiError(
        response,
        500,
        "UNSAFE_CATALOG_ASSET",
        "This catalog item references an unsafe asset path"
      );
      return;
    }

    const httpStatus = getHttpErrorStatus(error);
    if (httpStatus === 400 && error instanceof SyntaxError) {
      sendApiError(response, 400, "INVALID_JSON", "The request body contains invalid JSON");
      return;
    }
    if (httpStatus === 413) {
      sendApiError(response, 413, "REQUEST_TOO_LARGE", "The request body is too large");
      return;
    }

    log.error("Unhandled API failure", { requestId, path: request.path, error });
    sendApiError(response, 500, "INTERNAL_ERROR", "An unexpected server error occurred");
  };
  app.use(errorHandler);

  return app;
}
