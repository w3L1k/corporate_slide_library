# Final MVP checklist

Checked items are present in the repository or were verified by direct static/package inspection. External/manual and final command-run items stay unchecked until that evidence exists. Do not convert a pending check to `[x]` merely because the implementation looks correct.

## Repository and configuration

- [x] npm workspaces separate `apps/addin`, `apps/server`, and `packages/shared`.
- [x] Shared API/domain types and Zod schemas live in `packages/shared`.
- [x] Node.js 20+ is declared and a root lockfile is present.
- [x] `.env.example` documents loopback API, demo fallback, CORS, admin refresh, and frontend API settings without a secret.
- [x] Empty/unset `SLIDE_LIBRARY_PATH` resolves to the bundled `data/` library.
- [x] `HOST` defaults to `127.0.0.1`; wider network binding requires an explicit setting.
- [x] Root scripts cover browser/HTTPS development, build, lint, typecheck, tests, library/manifest validation, sideload, and content tools.
- [x] README and architecture, limitations, roadmap, demo, content-workflow, and final-checklist documents exist.

## Catalog and content

- [x] `data/catalog.json` contains 12 realistic metadata items across categories and statuses.
- [x] Every demo item has a registered PPTX and preview.
- [x] Package inspection confirms each of the 12 demo PPTX files contains one slide XML part.
- [x] Version, `updatedAt`, status, category, tags, owner/author/department/language fields are represented in the demo catalog.
- [x] The shared schema permits an empty catalog and strictly validates nonempty items.
- [x] IDs, statuses, timestamps, asset extensions, and safe relative path segments are schema-validated.
- [x] Full library validation checks duplicate IDs, registered-file presence/type, and resolved containment.
- [x] Optional Windows PowerPoint importer produces one-slide PPTX files, 1280×720 previews, and reviewable draft metadata without auto-publishing; a Windows COM smoke test passed.
- [x] Content publication, refresh semantics, and rollback are documented honestly.

## Backend and storage

- [x] Health, list, item, preview, and PPTX endpoints are implemented.
- [x] Personal asset list, upload, and registered-ID binary endpoints are implemented.
- [x] Personal deletion accepts only a validated registered UUID and removes both metadata and binary.
- [x] List API combines text, category, and status filters and returns total plus available categories.
- [x] Search covers title, description, category, tags, and optional `searchText`.
- [x] Unknown/invalid IDs and missing assets return stable friendly API errors.
- [x] Malformed and oversized JSON bodies map to client errors rather than generic 500 responses.
- [x] Catalog-invalid and catalog-unavailable failures are distinguished.
- [x] Binary routes resolve only a validated ID's registered source/preview; no client filesystem path is accepted.
- [x] Lexical containment and real-path checks reject traversal and link/junction escape.
- [x] Catalog cache reloads after a normal modification-time/size change; explicit `refresh()` bypasses the signature.
- [x] The pilot admin reindex route is absent by default and gated by configuration.
- [x] Server logs include startup/catalog/request/failure context and exclude PPTX/Base64 bodies.
- [x] Backend/API/storage/config/CLI automated tests are present.

## Task-pane UX

- [x] Catalog cards display preview, title, category, status, and updated date.
- [x] The sidebar can collapse to an icon rail and persists the user's choice locally.
- [x] User-facing controls, states, statuses, and errors are consistently localized in Russian; unfinished navigation placeholders are hidden.
- [x] Successfully inserted slide IDs are saved locally and exposed through a filterable, persistent Recent section.
- [x] Search is debounced and has a clear control.
- [x] Category/status filters, result count, and reset control are implemented.
- [x] Enlarged details dialog includes version and available governance metadata.
- [x] Dialog focus entry/return, Escape close, and Tab focus trapping are implemented.
- [x] Loading skeleton, API error/retry, empty-catalog, filtered-no-results, and preview-placeholder states are implemented.
- [x] Insertion disables competing insert controls and reports progress, success, unavailable host, and failure.
- [x] Personal PPTX, photo, illustration, and logo sections are clickable and their cards expose insertion controls with progress and browser fallback states.
- [x] Personal cards require an explicit confirmation before deletion and support cancellation.
- [x] Browser mode uses an explicit catalog-preview notice and never fakes insertion.
- [x] React UI and service automated tests are present.
- [ ] `npm run dev:browser` has been manually smoke-tested in the final revision with no unexpected console/network error.

## Office.js and manifest

- [x] Add-in-only XML manifest targets PowerPoint, requests `ReadWriteDocument`, exposes a Home-ribbon task-pane command, and declares `PowerPointApi 1.2`.
- [x] Office integration is behind `PowerPointService`; browser development uses a separate implementation.
- [x] Runtime initialization checks Office/PowerPoint globals, host readiness, and `PowerPointApi 1.2` support.
- [x] PPTX download is deferred until Insert.
- [x] Binary-to-Base64 conversion is chunked and covered by boundary/large-data tests.
- [x] `OfficePowerPointService` calls `PowerPoint.run`, `insertSlidesFromBase64`, `KeepSourceFormatting`, and `context.sync`.
- [x] Raster images use `ImageCoercion 1.1`; sanitized SVG uses `ImageCoercion 1.2`, both behind `PowerPointService`.
- [x] Unit tests verify Office insertion orchestration and unsupported/browser fallbacks through test doubles.
- [x] `npm run validate-manifest` passes in the final revision.
- [ ] `npm run sideload` successfully installs/opens the add-in in desktop PowerPoint.
- [ ] The **Open Slide Library** ribbon command opens the live task pane.
- [ ] The installed host confirms `PowerPointApi 1.2` and shows no initialization error.
- [ ] A demo item inserts exactly one editable slide into the current presentation with expected source formatting.
- [ ] Live insertion success and failure states have been observed in PowerPoint.
- [ ] `npm run sideload:stop` cleanly ends/removes the developer session.

## Quality gates

- [x] `npm run check` is defined to run library validation, manifest validation, lint, all typechecks, tests, and builds.
- [x] GitHub Actions runs `npm ci` and the complete `npm run check` gate for pushes, pull requests, and manual dispatches.
- [x] Tools have their own TypeScript project and are included in the root typecheck.
- [x] Automated coverage includes API filters/errors/CORS/path safety/refresh, task-pane UX, Base64 conversion, and PowerPoint service behavior.
- [x] `npm run validate-library` passes against all 12 final demo items.
- [x] `npm run lint` passes in the final revision.
- [x] `npm run typecheck` passes in the final revision, including tools.
- [x] `npm run test` passes: 53 server tests and 41 add-in tests.
- [x] `npm run build` passes in the final revision.
- [x] The aggregate `npm run check` passes: library validation, manifest validation, lint (0 errors), typecheck, 94 tests (53 server + 41 add-in), builds.
- [x] `npm audit` reports 0 known vulnerabilities for the installed dependency tree.

## Pilot readiness boundary

- [x] Known scope limits and pending verification are separated from confirmed bugs in `MVP_LIMITATIONS.md`.
- [x] Browser fallback and a 2–4 minute live demo are documented.
- [x] SharePoint replacement preserves the ID-based API, `SlideStorage`, and Office boundary in the architecture plan.
- [ ] A named operator, content owner, rollback copy, and supported PowerPoint build are assigned for a real department pilot.
- [ ] Identity, authorization, managed HTTPS hosting, and centrally managed deployment are implemented before exposing the service beyond a controlled local demo.

## Final sign-off record

Fill this only after the unchecked acceptance checks are run:

```text
Date/time:
Commit/revision:
Node/npm versions:
PowerPoint platform/version/build:
npm run check result:
Browser smoke result:
Sideload result:
Live insert result:
Known failures / evidence links:
Operator:
```
