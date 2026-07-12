import dotenv from "dotenv";
import path from "node:path";
import { fileURLToPath } from "node:url";

const REPOSITORY_ROOT = fileURLToPath(new URL("../../..", import.meta.url));

dotenv.config({ path: path.join(REPOSITORY_ROOT, ".env"), quiet: true });

const DEFAULT_DEMO_LIBRARY_PATH = path.join(REPOSITORY_ROOT, "data");

const DEFAULT_CORS_ORIGINS = [
  "https://localhost:3000",
  "http://localhost:3000",
  "http://127.0.0.1:3000"
];

export interface ServerConfig {
  host: string;
  port: number;
  libraryPath: string;
  corsOrigins: string[];
  enableAdminReindex: boolean;
}

function containsControlCharacter(value: string): boolean {
  return [...value].some((character) => {
    const codePoint = character.codePointAt(0);
    return codePoint !== undefined && (codePoint < 32 || codePoint === 127);
  });
}

function parseHost(rawValue: string | undefined): string {
  const host = rawValue?.trim() || "127.0.0.1";

  if (/\s/u.test(host) || containsControlCharacter(host)) {
    throw new Error("HOST must be a hostname or IP address without whitespace");
  }

  return host;
}

function parsePort(rawValue: string | undefined): number {
  const value = rawValue?.trim() || "3001";
  const port = Number(value);

  if (!Number.isInteger(port) || port < 1 || port > 65_535) {
    throw new Error(`PORT must be an integer between 1 and 65535; received ${value}`);
  }

  return port;
}

function parseCorsOrigins(rawValue: string | undefined): string[] {
  if (!rawValue?.trim()) {
    return [...DEFAULT_CORS_ORIGINS];
  }

  return [...new Set(rawValue.split(",").map((value) => value.trim()).filter(Boolean))];
}

function parseBoolean(rawValue: string | undefined): boolean {
  return rawValue?.trim().toLowerCase() === "true";
}

export function loadConfig(
  environment: NodeJS.ProcessEnv = process.env
): ServerConfig {
  const configuredLibraryPath = environment.SLIDE_LIBRARY_PATH?.trim();

  return {
    host: parseHost(environment.HOST),
    port: parsePort(environment.PORT),
    libraryPath: path.resolve(configuredLibraryPath || DEFAULT_DEMO_LIBRARY_PATH),
    corsOrigins: parseCorsOrigins(environment.CORS_ORIGINS),
    enableAdminReindex: parseBoolean(environment.ENABLE_ADMIN_REINDEX)
  };
}

export { DEFAULT_DEMO_LIBRARY_PATH };
