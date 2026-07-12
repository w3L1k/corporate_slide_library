export type LogFields = Record<string, unknown>;

export interface Logger {
  info(message: string, fields?: LogFields): void;
  warn(message: string, fields?: LogFields): void;
  error(message: string, fields?: LogFields): void;
}

function serializeError(value: unknown): unknown {
  if (!(value instanceof Error)) {
    return value;
  }

  return {
    name: value.name,
    message: value.message,
    stack: process.env.NODE_ENV === "development" ? value.stack : undefined
  };
}

function write(
  level: "info" | "warn" | "error",
  message: string,
  fields: LogFields = {}
): void {
  const safeFields = Object.fromEntries(
    Object.entries(fields).map(([key, value]) => [key, serializeError(value)])
  );
  const line = JSON.stringify({
    ...safeFields,
    timestamp: new Date().toISOString(),
    level,
    message
  });

  if (level === "error") {
    console.error(line);
  } else if (level === "warn") {
    console.warn(line);
  } else {
    console.info(line);
  }
}

export const logger: Logger = {
  info: (message, fields) => write("info", message, fields),
  warn: (message, fields) => write("warn", message, fields),
  error: (message, fields) => write("error", message, fields)
};
