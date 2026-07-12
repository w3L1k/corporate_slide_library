import type {
  ApiErrorResponse,
  SlideLibraryItem,
  SlideListResponse,
  SlideStatus
} from "@slide-library/shared";

export interface SlideListFilters {
  query?: string;
  category?: string;
  status?: SlideStatus;
}

export interface SlideLibraryApi {
  listSlides(filters: SlideListFilters, signal?: AbortSignal): Promise<SlideListResponse>;
  getSlide(id: string, signal?: AbortSignal): Promise<SlideLibraryItem>;
  downloadSlide(id: string, signal?: AbortSignal): Promise<ArrayBuffer>;
  getPreviewUrl(id: string): string;
}

export class ApiError extends Error {
  readonly status: number;
  readonly code: string | undefined;

  constructor(message: string, status: number, code?: string) {
    super(message);
    this.name = "ApiError";
    this.status = status;
    this.code = code;
  }
}

type FetchImplementation = typeof fetch;

const trimTrailingSlash = (value: string): string => value.replace(/\/+$/, "");

const isApiErrorResponse = (value: unknown): value is ApiErrorResponse => {
  if (typeof value !== "object" || value === null || !("error" in value)) {
    return false;
  }

  const error = value.error;
  return (
    typeof error === "object" &&
    error !== null &&
    "message" in error &&
    typeof error.message === "string" &&
    "code" in error &&
    typeof error.code === "string"
  );
};

const readError = async (response: Response): Promise<ApiError> => {
  let payload: unknown;

  try {
    payload = await response.json();
  } catch {
    payload = undefined;
  }

  if (isApiErrorResponse(payload)) {
    return new ApiError(payload.error.message, response.status, payload.error.code);
  }

  return new ApiError(
    response.status >= 500
      ? "The slide library is temporarily unavailable."
      : "The slide library request could not be completed.",
    response.status
  );
};

const requireOk = async (response: Response): Promise<Response> => {
  if (!response.ok) {
    throw await readError(response);
  }

  return response;
};

const isSlideListResponse = (value: unknown): value is SlideListResponse => {
  if (typeof value !== "object" || value === null) {
    return false;
  }

  return (
    "items" in value &&
    Array.isArray(value.items) &&
    "total" in value &&
    typeof value.total === "number" &&
    "availableCategories" in value &&
    Array.isArray(value.availableCategories) &&
    value.availableCategories.every((category) => typeof category === "string")
  );
};

export class HttpSlideLibraryApi implements SlideLibraryApi {
  private readonly baseUrl: string;
  private readonly fetchImplementation: FetchImplementation;

  constructor(
    baseUrl = "/api",
    fetchImplementation: FetchImplementation = globalThis.fetch.bind(globalThis)
  ) {
    this.baseUrl = trimTrailingSlash(baseUrl);
    this.fetchImplementation = fetchImplementation;
  }

  async listSlides(filters: SlideListFilters, signal?: AbortSignal): Promise<SlideListResponse> {
    const searchParams = new URLSearchParams();
    const query = filters.query?.trim();

    if (query) {
      searchParams.set("q", query);
    }
    if (filters.category) {
      searchParams.set("category", filters.category);
    }
    if (filters.status) {
      searchParams.set("status", filters.status);
    }

    const queryString = searchParams.toString();
    const response = await requireOk(
      await this.fetchImplementation(`${this.baseUrl}/slides${queryString ? `?${queryString}` : ""}`, {
        method: "GET",
        headers: { Accept: "application/json" },
        ...(signal ? { signal } : {})
      })
    );
    const payload: unknown = await response.json();

    if (!isSlideListResponse(payload)) {
      throw new ApiError("The slide library returned an invalid catalog response.", response.status);
    }

    return payload;
  }

  async getSlide(id: string, signal?: AbortSignal): Promise<SlideLibraryItem> {
    const response = await requireOk(
      await this.fetchImplementation(`${this.baseUrl}/slides/${encodeURIComponent(id)}`, {
        method: "GET",
        headers: { Accept: "application/json" },
        ...(signal ? { signal } : {})
      })
    );

    return (await response.json()) as SlideLibraryItem;
  }

  async downloadSlide(id: string, signal?: AbortSignal): Promise<ArrayBuffer> {
    const response = await requireOk(
      await this.fetchImplementation(`${this.baseUrl}/slides/${encodeURIComponent(id)}/file`, {
        method: "GET",
        headers: {
          Accept: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        },
        ...(signal ? { signal } : {})
      })
    );

    return response.arrayBuffer();
  }

  getPreviewUrl(id: string): string {
    return `${this.baseUrl}/slides/${encodeURIComponent(id)}/preview`;
  }
}

const configuredApiBaseUrl = import.meta.env.VITE_API_BASE_URL?.trim();

export const slideLibraryApi = new HttpSlideLibraryApi(configuredApiBaseUrl || "/api");
