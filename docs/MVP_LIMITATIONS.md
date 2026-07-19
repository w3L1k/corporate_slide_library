# MVP limitations and known bugs

## How this document uses the terms

- **MVP limitation**: an intentional scope or architecture boundary. The implemented behavior is working as designed, but it is not sufficient for a production or organization-wide rollout.
- **Bug**: implemented behavior that reproducibly contradicts the documented contract.
- **Pending verification**: code exists, but the required external host/environment has not yet been exercised. A pending check is not labeled a bug without a reproducible failure.

## Known bugs

No reproducible application defect is intentionally accepted in the current documented state.

Live desktop PowerPoint insertion and the automated sideload session remain explicit manual checklist items until they are exercised in a compatible installed host. That is a verification gap, not evidence that insertion works and not evidence of a bug. If the smoke test fails, record the PowerPoint platform, version/build, add-in error, and network response before reclassifying it here.

## Pending external verification

- Sideload `apps/addin/manifest.xml` into desktop PowerPoint through `npm run sideload`.
- Open the task pane from its ribbon command.
- Confirm that the host reports `PowerPointApi 1.2` support.
- Insert a demo item and confirm that exactly one editable slide appears with source formatting.
- Confirm the success/error toast and inspect the task-pane console/network log for unexpected errors.
- Stop and remove the developer sideload session with `npm run sideload:stop`.

The current state is tracked in [FINAL_CHECKLIST.md](FINAL_CHECKLIST.md).

## Intentional MVP limitations

### Storage and content governance

- The primary adapter is a local filesystem folder. There is no SharePoint, OneDrive, object-storage, database, or remote content connector.
- Content publication is file-based and manual. There is no admin UI, approval workflow, scheduled ingestion, content-owner notification, or automatic deprecation policy.
- The validator checks JSON/schema validity, unique IDs, registered-file presence, regular-file type, extension, and real-path containment. It does **not** open the PPTX, validate the Open XML package, count slides, scan macros/external links, or compare the preview with the source.
- “One item = one-slide PPTX” is therefore a required content-owner check. If a source file contains several slides, Office.js inserts all of them because the implementation does not pass `sourceSlideIds`.
- Preview generation is manual except for the optional Windows importer. The server does not render PPTX files.
- The importer depends on installed desktop PowerPoint and Windows COM automation. It writes draft metadata to `catalog.imported.json` for review; it does not merge or publish metadata automatically.
- A publish touches multiple files and is not transactional. There is no lock, staging service, immutable release, built-in backup, or rollback history.
- Replacing an asset under an existing registered path takes effect on its next binary request, even if catalog metadata has not changed. Content owners must keep the PPTX, preview, version, status, and `updatedAt` coherent.

### Catalog refresh and indexing

- Normal automatic reload compares only the catalog modification time and byte size. A tool that overwrites `catalog.json` while deliberately preserving both values can leave the cached metadata unchanged until a forced refresh or server restart.
- The task-pane **Refresh** button makes a new list request; it does not itself call the gated admin endpoint. A normal catalog edit is still detected by the storage signature check.
- `npm run reindex-library` validates and builds an index in a separate short-lived CLI process. It does not notify an already-running API process and does not write an index file.
- `POST /api/admin/reindex` performs a true forced refresh of the running process, but is absent unless `ENABLE_ADMIN_REINDEX=true`.

### Security and deployment

- There is no enterprise SSO, user identity, RBAC, per-department authorization, or content entitlement.
- The API listens on loopback by default, which limits exposure but is not authentication. Any local process that can reach it can request registered content.
- CORS restricts cooperating browsers; it does not authenticate clients or protect the API from non-browser callers.
- The optional admin reindex endpoint has no authentication. Enabling it is appropriate only in a controlled local pilot.
- The supplied manifest and icons point to `https://localhost:3000` and use a trusted development certificate. This is a development/sideload package, not a centrally deployable production manifest.
- Network-share sideloading is a Windows test mechanism, not a production distribution strategy. A pilot needs a real HTTPS origin and managed add-in deployment.
- There is no malware scanning, data-loss prevention integration, audit retention policy, rate limiting, or tenant-level compliance control.
- Personal asset deletion is permanent: the local MVP has no recycle bin, retention policy, or version history.

### PowerPoint behavior

- Real insertion requires a host/build that supports `PowerPointApi 1.2`; availability varies by platform and version. The manifest declares that minimum and the runtime also checks it.
- Personal raster-image insertion additionally requires `ImageCoercion 1.1`; SVG insertion requires `ImageCoercion 1.2`. The runtime checks these capabilities without disabling PPTX insertion on older compatible hosts.
- Browser development mode deliberately cannot modify a deck. It shows the catalog and an explicit unavailable message instead of simulating success.
- The service requests `KeepSourceFormatting`. There is no user option to use the destination theme.
- No `targetSlideId` is supplied. Per the Office.js option contract, the inserted slide uses the API default placement at the beginning of the presentation; there is no “insert after selected slide” control.
- The whole PPTX is downloaded into server memory, then browser memory, then represented as Base64 for Office.js. This is reasonable for one-slide pilot assets, not arbitrarily large decks.
- The MVP has not established a supported-platform test matrix across every Windows/Mac/web build. The official [PowerPoint requirement-set matrix](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets) is the compatibility source of truth.

### API, search, and product scope

- Search and filtering scan the in-memory catalog. There is no full-text engine, stemming, typo tolerance, synonym management, ranking, or pagination.
- The intended scale is tens to a few hundred items in one department, not a multi-tenant organization-wide catalog.
- There are no favorites, recently used items, recommendations, collections, usage analytics, insertion analytics, or offline mode.
- There is no version history in the product. `version` and `updatedAt` are owner-maintained metadata fields, not automatically generated revision records.
- There is one library root per server process and no UI for switching departments or libraries.
- The task-pane interface is localized in Russian. Demo slide titles, descriptions, and governance metadata intentionally remain in their source language, and there is no runtime language switch.
- Health reports that the HTTP process is responsive. It is not a deep health check of every registered asset or the PowerPoint host.

### Operations and reliability

- The API is one local Node.js process with an in-memory metadata cache. There is no clustering, high availability, job queue, distributed cache, or service-level objective.
- Structured logs go to standard output/error. There is no centralized log shipping, alerting, metrics backend, or privacy-retention policy.
- Backup, restore, and disaster recovery are external filesystem responsibilities.
- Dependency versions are installed from the repository lockfile, but dependency upgrades and Office client updates still require explicit compatibility testing before a wider pilot.

## Deferred work is not a bug

SharePoint storage, Entra ID, RBAC, centrally managed deployment, an admin/approval workflow, deep PPTX validation, analytics, and organization-scale search are staged in [PILOT_ROADMAP.md](PILOT_ROADMAP.md). Their absence should be evaluated as scope, while failures in the documented current flows should be reported as bugs with reproducible evidence.
