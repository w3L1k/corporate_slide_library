import { fileURLToPath } from "node:url";
import type { ServerOptions as HttpsServerOptions } from "node:https";
import { defineConfig, loadEnv } from "vite";
import react from "@vitejs/plugin-react";
import { getHttpsServerOptions } from "office-addin-dev-certs";

const workspaceRoot = fileURLToPath(new URL("../../", import.meta.url));

export default defineConfig(async ({ command, mode }) => {
  const env = loadEnv(mode, workspaceRoot, "");
  const browserMode = (process.env.ADDIN_BROWSER_MODE ?? env.ADDIN_BROWSER_MODE) === "true";
  const https =
    command === "serve" && !browserMode
      ? (await getHttpsServerOptions()) as HttpsServerOptions
      : undefined;
  const apiTarget =
    (process.env.VITE_API_PROXY_TARGET ?? env.VITE_API_PROXY_TARGET)?.trim() ||
    "http://127.0.0.1:3001";

  return {
    envDir: workspaceRoot,
    plugins: react(),
    server: {
      host: "0.0.0.0",
      port: 3000,
      strictPort: true,
      ...(https ? { https } : {}),
      proxy: {
        "/api": {
          target: apiTarget,
          changeOrigin: true,
          secure: false
        }
      }
    },
    preview: {
      host: "0.0.0.0",
      port: 3000,
      strictPort: true
    }
  };
});
