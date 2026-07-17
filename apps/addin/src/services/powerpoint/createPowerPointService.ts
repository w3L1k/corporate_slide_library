import type { SlideLibraryApi } from "../api";
import { BrowserPowerPointService } from "./BrowserPowerPointService";
import {
  OfficePowerPointService,
  type ImageCoercionType,
  type OfficeImageInsertionRuntime,
  type PowerPointRequestContextRuntime,
  type PowerPointRuntime
} from "./OfficePowerPointService";
import type { PowerPointService } from "./types";

const POWERPOINT_HOST_NAME = "PowerPoint";
const POWERPOINT_API_VERSION = "1.2";
const IMAGE_COERCION_VERSION = "1.1";
const SVG_COERCION_VERSION = "1.2";
const DEFAULT_READY_TIMEOUT_MS = 1_500;
const UNSUPPORTED_POWERPOINT_MESSAGE =
  "This version of PowerPoint does not support inserting slides from the library. Update PowerPoint and try again.";
const POWERPOINT_INITIALIZATION_MESSAGE =
  "PowerPoint integration could not be initialized. Close and reopen the add-in, then try again.";

export interface OfficeReadyInfoRuntime {
  host?: string | null;
}

export interface OfficeAsyncResultRuntime {
  status: unknown;
  error?: {
    message?: string;
  } | null;
}

export interface OfficeDocumentRuntime {
  setSelectedDataAsync(
    data: string,
    options: { coercionType: ImageCoercionType },
    callback: (result: OfficeAsyncResultRuntime) => void
  ): void;
}

export interface OfficeRequirementsRuntime {
  isSetSupported(name: string, minVersion?: string): boolean;
}

export interface OfficeRuntime {
  onReady(): Promise<OfficeReadyInfoRuntime>;
  AsyncResultStatus?: {
    Succeeded: unknown;
  };
  context?: {
    requirements?: OfficeRequirementsRuntime;
    document?: OfficeDocumentRuntime;
  };
}

export interface PowerPointGlobals {
  office?: OfficeRuntime;
  powerPoint?: PowerPointRuntime;
}

const getGlobals = (): PowerPointGlobals => {
  const runtime = globalThis as typeof globalThis & {
    Office?: OfficeRuntime;
    PowerPoint?: {
      run<T>(
        batch: (context: PowerPoint.RequestContext) => Promise<T>
      ): Promise<T>;
    };
  };

  const powerPoint = runtime.PowerPoint
    ? {
        run: <T>(batch: (context: PowerPointRequestContextRuntime) => Promise<T>) =>
          runtime.PowerPoint!.run((context) =>
            batch(context as unknown as PowerPointRequestContextRuntime)
          )
      }
    : undefined;

  return {
    ...(runtime.Office ? { office: runtime.Office } : {}),
    ...(powerPoint ? { powerPoint } : {})
  };
};

const withTimeout = async <T>(promise: Promise<T>, timeoutMs: number): Promise<T | undefined> => {
  let timeoutHandle: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<undefined>((resolve) => {
    timeoutHandle = setTimeout(() => resolve(undefined), timeoutMs);
  });

  try {
    return await Promise.race([promise, timeout]);
  } finally {
    if (timeoutHandle !== undefined) {
      clearTimeout(timeoutHandle);
    }
  }
};

const supportsRequirement = (
  requirements: OfficeRequirementsRuntime | undefined,
  name: string,
  version: string
): boolean => {
  try {
    return requirements?.isSetSupported(name, version) ?? false;
  } catch {
    return false;
  }
};

const createImageInsertionRuntime = (
  office: OfficeRuntime
): OfficeImageInsertionRuntime | undefined => {
  const document = office.context?.document;
  const succeededStatus = office.AsyncResultStatus?.Succeeded;
  if (!document || succeededStatus === undefined) {
    return undefined;
  }

  return {
    insert: (data, coercionType) =>
      new Promise<void>((resolve, reject) => {
        document.setSelectedDataAsync(data, { coercionType }, (result) => {
          if (result.status === succeededStatus) {
            resolve();
            return;
          }

          reject(
            new Error(
              result.error?.message?.trim() || "PowerPoint could not insert the selected image."
            )
          );
        });
      })
  };
};

export const createPowerPointService = async (
  api: Pick<SlideLibraryApi, "downloadSlide" | "downloadPersonalAsset">,
  providedGlobals?: PowerPointGlobals,
  readyTimeoutMs = DEFAULT_READY_TIMEOUT_MS
): Promise<PowerPointService> => {
  const initialGlobals = providedGlobals ?? getGlobals();

  // In a browser there is no Office global, so do not wait. Inside an Office host,
  // PowerPoint can be exposed only after Office.onReady resolves.
  if (!initialGlobals.office) {
    return new BrowserPowerPointService();
  }

  let readyInfo: OfficeReadyInfoRuntime | undefined;

  try {
    readyInfo = await withTimeout(initialGlobals.office.onReady(), readyTimeoutMs);
  } catch {
    return new BrowserPowerPointService(POWERPOINT_INITIALIZATION_MESSAGE);
  }

  if (!readyInfo) {
    return new BrowserPowerPointService();
  }

  if (readyInfo.host !== POWERPOINT_HOST_NAME) {
    return new BrowserPowerPointService();
  }

  const readyGlobals = providedGlobals ?? getGlobals();
  const readyOffice = readyGlobals.office ?? initialGlobals.office;
  const powerPoint = readyGlobals.powerPoint;
  if (!powerPoint) {
    return new BrowserPowerPointService(POWERPOINT_INITIALIZATION_MESSAGE);
  }

  const requirements = readyOffice.context?.requirements;
  const supportsSlideInsertion = supportsRequirement(
    requirements,
    "PowerPointApi",
    POWERPOINT_API_VERSION
  );

  if (!supportsSlideInsertion) {
    return new BrowserPowerPointService(UNSUPPORTED_POWERPOINT_MESSAGE);
  }

  const supportsImageInsertion = supportsRequirement(
    requirements,
    "ImageCoercion",
    IMAGE_COERCION_VERSION
  );
  const supportsSvgInsertion =
    supportsImageInsertion &&
    supportsRequirement(requirements, "ImageCoercion", SVG_COERCION_VERSION);
  const imageInsertionRuntime = supportsImageInsertion
    ? createImageInsertionRuntime(readyOffice)
    : undefined;

  return new OfficePowerPointService(api, powerPoint, {
    ...(imageInsertionRuntime ? { imageInsertionRuntime } : {}),
    supportsSvgInsertion
  });
};
