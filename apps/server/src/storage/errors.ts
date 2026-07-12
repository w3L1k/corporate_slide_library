export type AssetKind = "slide" | "preview";

export class CatalogReadError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = "CatalogReadError";
  }
}

export class CatalogValidationError extends Error {
  constructor(
    message: string,
    public readonly issues: string[] = []
  ) {
    super(message);
    this.name = "CatalogValidationError";
  }
}

export class SlideNotFoundError extends Error {
  constructor(public readonly id: string) {
    super(`Slide '${id}' is not registered in the catalog`);
    this.name = "SlideNotFoundError";
  }
}

export class AssetNotFoundError extends Error {
  constructor(
    public readonly id: string,
    public readonly kind: AssetKind,
    options?: ErrorOptions
  ) {
    super(`The registered ${kind} asset for slide '${id}' is unavailable`, options);
    this.name = "AssetNotFoundError";
  }
}

export class UnsafeAssetPathError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "UnsafeAssetPathError";
  }
}
