# Technical plan and architecture decisions

## Delivery plan

1. Establish an npm-workspaces repository with a React task pane, Express API, and shared TypeScript contract.
2. Implement a validated filesystem storage adapter, searchable catalog API, binary preview/PPTX routes, and content tooling.
3. Generate a realistic demo library of one-slide PPTX files with matching previews.
4. Build the responsive task-pane experience with loading, error, empty, detail, refresh, and insertion states.
5. Integrate Office.js through a dedicated service and insert downloaded presentations with `insertSlidesFromBase64`.
6. Verify automated quality gates, browser mode, manifest validity, and the PowerPoint sideload path.

## Architecture decisions

- **npm workspaces, not microservices.** One repository keeps the pilot easy to install while retaining explicit add-in, server, and shared-package boundaries.
- **JSON catalog + filesystem storage.** This is inspectable and dependable for a department pilot; a `SlideStorage` interface isolates a later SharePoint or corporate API adapter.
- **One-slide PPTX per item.** PowerPoint can insert the source slide itself, preserving editable content, layouts, charts, and notes instead of flattening it to an image.
- **Same-origin API in development.** The HTTPS Vite server proxies `/api` to the local HTTP backend, avoiding mixed-content failures inside the Office task pane.
- **Validated registered paths only.** The server resolves files from catalog metadata after schema, extension, and containment checks. Route parameters never become filesystem paths.
- **Office integration behind an interface.** `OfficePowerPointService` owns Office.js calls; `BrowserPowerPointService` provides a deliberate, user-facing development-mode response.
- **No database or enterprise authentication in the MVP.** Both belong at the pilot expansion boundary, not in the local demonstration path.
