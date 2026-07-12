import type { SlideLibraryApi } from "../api";
import { BrowserPowerPointService } from "./BrowserPowerPointService";
import {
  OfficePowerPointService,
  type PowerPointRequestContextRuntime,
  type PowerPointRuntime
} from "./OfficePowerPointService";
import type { PowerPointService } from "./types";

const POWERPOINT_HOST_NAME = "PowerPoint";
const POWERPOINT_API_VERSION = "1.2";
const DEFAULT_READY_TIMEOUT_MS = 1_500;
const UNSUPPORTED_POWERPOINT_MESSAGE =
  "This version of PowerPoint does not support inserting slides from the library. Update PowerPoint and try again.";
const POWERPOINT_INITIALIZATION_MESSAGE =
  "PowerPoint integration could not be initialized. Close and reopen the add-in, then try again.";

export interface OfficeReadyInfoRuntime {
  host?: string | null;
}

export interface OfficeRuntime {
  onReady(): Promise<OfficeReadyInfoRuntime>;
  context?: {
    requirements?: {
      isSetSupported(name: string, minVersion?: string): boolean;
    };
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

export const createPowerPointService = async (
  api: Pick<SlideLibraryApi, "downloadSlide">,
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
  const powerPoint = readyGlobals.powerPoint;
  if (!powerPoint) {
    return new BrowserPowerPointService(POWERPOINT_INITIALIZATION_MESSAGE);
  }

  const requirements = initialGlobals.office.context?.requirements;
  let supportsSlideInsertion: boolean;

  try {
    supportsSlideInsertion =
      requirements?.isSetSupported("PowerPointApi", POWERPOINT_API_VERSION) ?? false;
  } catch {
    supportsSlideInsertion = false;
  }

  if (!supportsSlideInsertion) {
    return new BrowserPowerPointService(UNSUPPORTED_POWERPOINT_MESSAGE);
  }

  return new OfficePowerPointService(api, powerPoint);
};
