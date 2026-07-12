import { z } from "zod";

export const slideStatusValues = ["approved", "draft", "deprecated"] as const;
export const SlideStatusSchema = z.enum(slideStatusValues);
export type SlideStatus = z.infer<typeof SlideStatusSchema>;

const relativeAssetPath = z
  .string()
  .min(1)
  .max(260)
  .refine((value) => !value.includes("\\"), "Use forward slashes in catalog paths")
  .refine((value) => !value.startsWith("/"), "Asset path must be relative")
  .refine((value) => !value.includes(":"), "Asset path cannot contain a drive or URI scheme")
  .refine(
    (value) =>
      ![...value].some((character) => {
        const codePoint = character.codePointAt(0) ?? 0;
        return codePoint < 32 || codePoint === 127;
      }),
    "Asset path cannot contain control characters"
  )
  .refine(
    (value) => !value.split("/").some((segment) => segment.length === 0 || segment === ".." || segment === "."),
    "Asset path cannot contain empty or dot segments"
  );

export const SlideLibraryItemSchema = z
  .object({
    id: z.string().regex(/^[a-z0-9]+(?:-[a-z0-9]+)*$/, "Use a URL-safe kebab-case id"),
    title: z.string().trim().min(1).max(120),
    description: z.string().trim().min(1).max(500).optional(),
    category: z.string().trim().min(1).max(80),
    tags: z.array(z.string().trim().min(1).max(50)).min(1).max(30),
    department: z.string().trim().min(1).max(80).optional(),
    language: z.string().trim().min(2).max(20).optional(),
    version: z.string().trim().min(1).max(30),
    status: SlideStatusSchema,
    updatedAt: z.iso.datetime({ offset: true }),
    sourceFile: relativeAssetPath.refine(
      (value) => value.toLowerCase().endsWith(".pptx"),
      "Source file must be a .pptx"
    ),
    previewFile: relativeAssetPath.refine(
      (value) => /\.(?:png|jpe?g|webp)$/i.test(value),
      "Preview must be PNG, JPEG, or WebP"
    ),
    author: z.string().trim().min(1).max(100).optional(),
    owner: z.string().trim().min(1).max(100).optional(),
    searchText: z.string().trim().max(1000).optional()
  })
  .strict();

export type SlideLibraryItem = z.infer<typeof SlideLibraryItemSchema>;

export const CatalogSchema = z.array(SlideLibraryItemSchema);
export type Catalog = z.infer<typeof CatalogSchema>;

export interface SlideListResponse {
  items: SlideLibraryItem[];
  total: number;
  availableCategories: string[];
}

export interface HealthResponse {
  status: "ok";
}

export interface ApiErrorResponse {
  error: {
    code: string;
    message: string;
  };
}

export interface ReindexResponse {
  status: "ok";
  itemCount: number;
  refreshedAt: string;
}

export interface LibraryValidationIssue {
  itemId?: string;
  path?: string;
  message: string;
}

export interface LibraryValidationReport {
  valid: boolean;
  itemCount: number;
  errors: LibraryValidationIssue[];
  warnings: LibraryValidationIssue[];
}
