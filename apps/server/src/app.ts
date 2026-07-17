import {
  personalAssetKindValues,
  slideStatusValues,
  type ApiErrorResponse,
  type HealthResponse,
  type PersonalAssetKind,
  type PersonalAssetListResponse,
  type ReindexResponse,
  type SlideLibraryItem,
  type SlideListResponse,
  type SlideStatus
} from "@slide-library/shared";
import cors from "cors";
import multer from "multer";
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
import type { PersonalAssetStorage } from "./storage/PersonalAssetStorage.js";
import {
  AssetNotFoundError,
  CatalogReadError,
  CatalogValidationError,
  SlideNotFoundError,
  UnsafeAssetPathError
} from "./storage/errors.js";

export interface CreateAppOptions {
  storage: SlideStorage;
  personalAssetStorage?: PersonalAssetStorage;
  config: Pick<ServerConfig, "corsOrigins" | "enableAdminReindex">;
  logger?: Logger;
}

const SLIDE_ID_PATTERN = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;
const POWERPOINT_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const UUID_PATTERN =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const MAX_PERSONAL_ASSET_SIZE = 20 * 1024 * 1024;

interface ValidatedUpload {
  extension: string;
  mimeType: string;
}

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

function requirePersonalAssetId(request: Request): string {
  const id = request.params.id;
  if (typeof id !== "string" || !UUID_PATTERN.test(id)) {
    throw new ApiError(400, "INVALID_PERSONAL_ASSET_ID", "Personal asset id has an invalid format");
  }
  return id;
}

function startsWithBytes(buffer: Buffer, expected: readonly number[]): boolean {
  return expected.every((value, index) => buffer[index] === value);
}

function isWebp(buffer: Buffer): boolean {
  return (
    buffer.length >= 12 &&
    buffer.subarray(0, 4).toString("ascii") === "RIFF" &&
    buffer.subarray(8, 12).toString("ascii") === "WEBP"
  );
}

function isSafeSvg(buffer: Buffer): boolean {
  const text = buffer.toString("utf8").trim();
  return (
    /^(?:<\?xml[^>]*>\s*)?<svg[\s>]/iu.test(text) &&
    !/<script[\s>]/iu.test(text) &&
    !/<foreignObject[\s>]/iu.test(text) &&
    !/\son[a-z]+\s*=/iu.test(text) &&
    !/javascript\s*:/iu.test(text)
  );
}

function validatePersonalUpload(
  kind: PersonalAssetKind,
  file: Express.Multer.File
): ValidatedUpload {
  const extension = path.extname(file.originalname).toLowerCase();
  const buffer = file.buffer;
  const isImageKind =
    kind === "photo" || kind === "illustration" || kind === "logo";

  if (
    kind === "presentation" &&
    extension === ".pptx" &&
    startsWithBytes(buffer, [0x50, 0x4b])
  ) {
    return { extension, mimeType: POWERPOINT_CONTENT_TYPE };
  }

  if (
    isImageKind &&
    extension === ".png" &&
    startsWithBytes(buffer, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  ) {
    return { extension, mimeType: "image/png" };
  }

  if (
    isImageKind &&
    [".jpg", ".jpeg"].includes(extension) &&
    startsWithBytes(buffer, [0xff, 0xd8, 0xff])
  ) {
    return { extension, mimeType: "image/jpeg" };
  }

  if (
    isImageKind &&
    extension === ".webp" &&
    isWebp(buffer)
  ) {
    return { extension, mimeType: "image/webp" };
  }

  if (
    (kind === "illustration" || kind === "logo") &&
    extension === ".svg" &&
    isSafeSvg(buffer)
  ) {
    return { extension, mimeType: "image/svg+xml" };
  }

  throw new ApiError(
    400,
    "UNSUPPORTED_PERSONAL_ASSET",
    "Upload a PPTX presentation, PNG/JPEG/WebP photo, or PNG/JPEG/WebP/SVG illustration or logo"
  );
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
  const { storage, personalAssetStorage, config } = options;
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
      methods: ["GET", "POST", "DELETE", "OPTIONS"],
      optionsSuccessStatus: 204
    })
  );
  app.use(express.json({ limit: "16kb" }));

  app.get("/api/health", (_request, response: Response<HealthResponse>) => {
    response.json({ status: "ok" });
  });

  if (personalAssetStorage) {
    const upload = multer({
      storage: multer.memoryStorage(),
      limits: { fileSize: MAX_PERSONAL_ASSET_SIZE, files: 1, fields: 4 }
    });

    app.get(
      "/api/personal-assets",
      asyncRoute(async (_request, response: Response<PersonalAssetListResponse>) => {
        const items = await personalAssetStorage.list();
        response.json({ items, total: items.length });
      })
    );

    app.post(
      "/api/personal-assets",
      upload.single("file"),
      asyncRoute(async (request, response) => {
        const rawKind = typeof request.body.kind === "string" ? request.body.kind : "";
        if (!personalAssetKindValues.includes(rawKind as PersonalAssetKind)) {
          throw new ApiError(
            400,
            "INVALID_PERSONAL_ASSET_KIND",
            `kind must be one of: ${personalAssetKindValues.join(", ")}`
          );
        }
        if (!request.file) {
          throw new ApiError(400, "PERSONAL_ASSET_FILE_REQUIRED", "Select a file to upload");
        }

        const title =
          typeof request.body.title === "string" ? request.body.title.trim() : "";
        if (!title || title.length > 120) {
          throw new ApiError(
            400,
            "INVALID_PERSONAL_ASSET_TITLE",
            "title must contain between 1 and 120 characters"
          );
        }

        const kind = rawKind as PersonalAssetKind;
        const validated = validatePersonalUpload(kind, request.file);
        const item = await personalAssetStorage.add({
          title,
          kind,
          originalFileName: path.basename(request.file.originalname).slice(0, 180),
          mimeType: validated.mimeType,
          extension: validated.extension,
          data: request.file.buffer
        });
        response.status(201).json(item);
      })
    );

    app.get(
      "/api/personal-assets/:id/file",
      asyncRoute(async (request, response) => {
        const id = requirePersonalAssetId(request);
        const asset = await personalAssetStorage.getFile(id);
        if (!asset) {
          throw new ApiError(404, "PERSONAL_ASSET_NOT_FOUND", "Personal asset was not found");
        }

        response.setHeader("Content-Type", asset.item.mimeType);
        response.setHeader("Content-Disposition", `inline; filename="asset-${asset.item.id}"`);
        response.setHeader("Cache-Control", "private, max-age=0, must-revalidate");
        if (asset.item.mimeType === "image/svg+xml") {
          response.setHeader("Content-Security-Policy", "default-src 'none'; sandbox");
        }
        response.send(asset.data);
      })
    );

    app.delete(
      "/api/personal-assets/:id",
      asyncRoute(async (request, response) => {
        const id = requirePersonalAssetId(request);
        const removed = await personalAssetStorage.remove(id);
        if (!removed) {
          throw new ApiError(404, "PERSONAL_ASSET_NOT_FOUND", "Personal asset was not found");
        }

        response.status(204).send();
      })
    );
  }

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
    if (error instanceof multer.MulterError) {
      sendApiError(
        response,
        error.code === "LIMIT_FILE_SIZE" ? 413 : 400,
        error.code === "LIMIT_FILE_SIZE" ? "PERSONAL_ASSET_TOO_LARGE" : "INVALID_UPLOAD",
        error.code === "LIMIT_FILE_SIZE"
          ? "Personal assets must not exceed 20 MB"
          : "The personal asset upload is invalid"
      );
      return;
    }

    log.error("Unhandled API failure", { requestId, path: request.path, error });
    sendApiError(response, 500, "INTERNAL_ERROR", "An unexpected server error occurred");
  };
  app.use(errorHandler);

  return app;
}
