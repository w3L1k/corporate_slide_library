import path from "node:path";
import { describe, expect, it } from "vitest";

import { DEFAULT_DEMO_LIBRARY_PATH, loadConfig } from "../src/config.js";

describe("server configuration", () => {
  it("uses a loopback-only server and the demo library by default", () => {
    const config = loadConfig({});

    expect(config.host).toBe("127.0.0.1");
    expect(config.port).toBe(3001);
    expect(config.libraryPath).toBe(DEFAULT_DEMO_LIBRARY_PATH);
    expect(config.enableAdminReindex).toBe(false);
    expect(config.corsOrigins).toContain("https://localhost:3000");
  });

  it("parses explicit pilot settings and removes duplicate CORS origins", () => {
    const config = loadConfig({
      HOST: " 0.0.0.0 ",
      PORT: "4310",
      SLIDE_LIBRARY_PATH: "./custom-library",
      CORS_ORIGINS: "https://pilot.example, https://pilot.example",
      ENABLE_ADMIN_REINDEX: "TRUE"
    });

    expect(config).toMatchObject({
      host: "0.0.0.0",
      port: 4310,
      libraryPath: path.resolve("./custom-library"),
      corsOrigins: ["https://pilot.example"],
      enableAdminReindex: true
    });
  });

  it.each([
    [{ PORT: "0" }, /PORT must be an integer/],
    [{ PORT: "12.5" }, /PORT must be an integer/],
    [{ HOST: "bad host" }, /HOST must be a hostname/]
  ])("rejects unsafe configuration: %o", (environment, expectedMessage) => {
    expect(() => loadConfig(environment)).toThrow(expectedMessage);
  });
});
