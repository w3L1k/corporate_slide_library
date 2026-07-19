import type {
  ApiErrorResponse,
  PersonalAsset,
  PersonalAssetKind,
  PersonalAssetListResponse,
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
  listPersonalAssets(signal?: AbortSignal): Promise<PersonalAssetListResponse>;
  uploadPersonalAsset(
    kind: PersonalAssetKind,
    title: string,
    file: File,
    signal?: AbortSignal
  ): Promise<PersonalAsset>;
  downloadPersonalAsset(id: string, signal?: AbortSignal): Promise<ArrayBuffer>;
  deletePersonalAsset(id: string, signal?: AbortSignal): Promise<void>;
  getPersonalAssetFileUrl(id: string): string;
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

const API_ERROR_MESSAGES: Readonly<Record<string, string>> = {
  CATALOG_INVALID: "Каталог повреждён или содержит некорректные данные.",
  CATALOG_UNAVAILABLE: "Каталог временно недоступен.",
  INVALID_PERSONAL_ASSET_ID: "Указан некорректный идентификатор личного материала.",
  INVALID_PERSONAL_ASSET_KIND: "Выбран неподдерживаемый тип личного материала.",
  INVALID_PERSONAL_ASSET_TITLE: "Проверьте название личного материала.",
  INVALID_QUERY: "Проверьте параметры поиска.",
  INVALID_SLIDE_ID: "Указан некорректный идентификатор слайда.",
  INVALID_STATUS: "Выбран неподдерживаемый статус.",
  PERSONAL_ASSET_FILE_REQUIRED: "Выберите файл для загрузки.",
  PERSONAL_ASSET_NOT_FOUND: "Личный материал не найден.",
  PERSONAL_ASSET_TOO_LARGE: "Размер личного материала превышает допустимый лимит.",
  PREVIEW_NOT_FOUND: "Предпросмотр слайда не найден.",
  REQUEST_TOO_LARGE: "Размер запроса превышает допустимый лимит.",
  SLIDE_FILE_NOT_FOUND: "Файл слайда не найден.",
  SLIDE_NOT_FOUND: "Слайд не найден.",
  UNSAFE_CATALOG_ASSET: "Материал каталога не прошёл проверку безопасности.",
  UNSUPPORTED_PERSONAL_ASSET: "Формат личного материала не поддерживается."
};

const getFallbackRequestError = (status: number): string =>
  status >= 500
    ? "Библиотека слайдов временно недоступна."
    : "Не удалось выполнить запрос к библиотеке слайдов.";

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
    return new ApiError(
      API_ERROR_MESSAGES[payload.error.code] ?? getFallbackRequestError(response.status),
      response.status,
      payload.error.code
    );
  }

  return new ApiError(getFallbackRequestError(response.status), response.status);
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

const isPersonalAssetListResponse = (value: unknown): value is PersonalAssetListResponse =>
  typeof value === "object" &&
  value !== null &&
  "items" in value &&
  Array.isArray(value.items) &&
  "total" in value &&
  typeof value.total === "number";

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
      throw new ApiError(
        "Сервер вернул некорректный каталог слайдов.",
        response.status
      );
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

  async listPersonalAssets(signal?: AbortSignal): Promise<PersonalAssetListResponse> {
    const response = await requireOk(
      await this.fetchImplementation(`${this.baseUrl}/personal-assets`, {
        method: "GET",
        headers: { Accept: "application/json" },
        ...(signal ? { signal } : {})
      })
    );
    const payload: unknown = await response.json();
    if (!isPersonalAssetListResponse(payload)) {
      throw new ApiError(
        "Сервер вернул некорректный список личных материалов.",
        response.status
      );
    }
    return payload;
  }

  async uploadPersonalAsset(
    kind: PersonalAssetKind,
    title: string,
    file: File,
    signal?: AbortSignal
  ): Promise<PersonalAsset> {
    const formData = new FormData();
    formData.set("kind", kind);
    formData.set("title", title);
    formData.set("file", file);

    const response = await requireOk(
      await this.fetchImplementation(`${this.baseUrl}/personal-assets`, {
        method: "POST",
        body: formData,
        ...(signal ? { signal } : {})
      })
    );
    return (await response.json()) as PersonalAsset;
  }

  async downloadPersonalAsset(id: string, signal?: AbortSignal): Promise<ArrayBuffer> {
    const response = await requireOk(
      await this.fetchImplementation(
        `${this.baseUrl}/personal-assets/${encodeURIComponent(id)}/file`,
        {
          method: "GET",
          headers: { Accept: "*/*" },
          ...(signal ? { signal } : {})
        }
      )
    );
    return response.arrayBuffer();
  }

  async deletePersonalAsset(id: string, signal?: AbortSignal): Promise<void> {
    await requireOk(
      await this.fetchImplementation(
        `${this.baseUrl}/personal-assets/${encodeURIComponent(id)}`,
        {
          method: "DELETE",
          ...(signal ? { signal } : {})
        }
      )
    );
  }

  getPersonalAssetFileUrl(id: string): string {
    return `${this.baseUrl}/personal-assets/${encodeURIComponent(id)}/file`;
  }
}

const configuredApiBaseUrl = import.meta.env.VITE_API_BASE_URL?.trim();

export const slideLibraryApi = new HttpSlideLibraryApi(configuredApiBaseUrl || "/api");
