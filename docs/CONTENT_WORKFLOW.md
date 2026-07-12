# Content workflow

This is the MVP operating procedure for content owners. There is no admin UI or transactional publisher, so validation and publication order are part of the safety model.

## Roles

- **Author** prepares or updates the PowerPoint slide and preview.
- **Content owner** verifies accuracy, ownership, version, date, and status.
- **Pilot operator** publishes the validated files and catalog, observes the API, and can roll back.

One person may hold all roles in a local demo, but the checks should still be performed explicitly.

## Library layout

```text
library-root/
  catalog.json
  slides/
    company-overview.pptx
  previews/
    company-overview.png
```

`catalog.json` is a JSON array. An empty array is valid. Every nonempty item has one registered source and one registered preview.

## Prepare the source slide

1. Start from approved content, remove every other slide, and save as `.pptx`.
2. Reopen the saved file and verify that it contains **exactly one slide**.
3. Check that text, shapes, charts, and tables remain editable.
4. Check fonts, theme/layout dependencies, linked media, external data, notes, and any sensitive hidden content.
5. Insert that file manually into a test deck with source formatting and inspect the result.
6. Export a representative preview as PNG, JPEG, or WebP. The demo uses 1280×720 PNG for a 16:9 slide; those dimensions are a convention, not a schema requirement.

The automated validator does not parse the Open XML package or count slides. These checks cannot be skipped merely because validation passes.

## Optional Windows bulk import

An existing multi-slide deck can be split through installed desktop PowerPoint on Windows:

```powershell
npm run import-pptx -- `
  -SourcePptx "C:\Input\corporate-deck.pptx" `
  -LibraryRoot "C:\CorporateSlideLibrary" `
  -Category "Imported" `
  -Owner "Content owner"
```

The script:

- uses PowerPoint COM automation and therefore requires Windows desktop PowerPoint;
- creates one `.pptx` per source slide under `slides/`;
- exports 1280×720 PNG previews under `previews/`;
- creates stable kebab-style IDs with the source slide index;
- marks every generated item `draft` at version `1.0`;
- writes generated metadata to `catalog.imported.json`.

It does **not** modify `catalog.json`. Review titles, descriptions, category, tags, owner, status, version, and paths; then merge only approved objects. Use `-Overwrite` only after backing up outputs, because it allows existing generated files with matching IDs to be replaced.

## Catalog item schema

Example:

```json
{
  "id": "revenue-overview",
  "title": "Revenue Overview",
  "description": "Quarterly revenue trend with management commentary.",
  "category": "Finance",
  "tags": ["revenue", "quarterly", "finance"],
  "department": "Finance",
  "language": "en-US",
  "version": "3.2",
  "status": "approved",
  "updatedAt": "2026-07-11T16:00:00+03:00",
  "sourceFile": "slides/revenue-overview-v3-2.pptx",
  "previewFile": "previews/revenue-overview-v3-2.png",
  "author": "FP&A",
  "owner": "Chief Financial Office",
  "searchText": "sales recurring revenue yoy"
}
```

Required fields are `id`, `title`, `category`, `tags`, `version`, `status`, `updatedAt`, `sourceFile`, and `previewFile`.

Important rules enforced by the shared schema:

- `id` is lowercase kebab-case (`a-z`, digits, and single hyphen-separated segments) and must be unique.
- `status` is exactly `approved`, `draft`, or `deprecated`.
- `tags` contains 1–30 nonempty strings.
- `updatedAt` is an ISO 8601 date/time with timezone information.
- `sourceFile` ends in `.pptx`; a preview ends in `.png`, `.jpg`, `.jpeg`, or `.webp`.
- Asset paths are relative to the library root, use forward slashes, and contain no empty, dot, traversal, drive, URI, backslash, or control-character segment.
- The item schema is strict; unknown fields are rejected.
- Text fields have bounded lengths. Run the validator instead of relying on this summary for every limit.

Use `searchText` for approved synonyms or business terminology that is not already in the visible fields. Search also covers title, description, category, and tags; it does not currently search owner or department.

## Validate a candidate library

With the bundled demo or `SLIDE_LIBRARY_PATH` from `.env`:

```powershell
npm run validate-library
```

With an explicit candidate root:

```powershell
npm run validate-library -- --path "C:\CorporateSlideLibrary"
```

A positional directory and `--path=<directory>` are also supported. CLI path precedence is: command argument, `SLIDE_LIBRARY_PATH`, then bundled `data/`.

The command emits a JSON report with the resolved `libraryPath`, `valid`, `itemCount`, `errors`, and `warnings`. It exits nonzero when invalid. Current full validation checks:

- readable JSON array and item schema;
- duplicate IDs;
- readable library root;
- source and preview existence;
- regular-file type;
- safe lexical and resolved containment, including symlink/junction escape detection.

It does not scan unregistered/orphan files and does not inspect PPTX internals. The `warnings` array is currently reserved and normally empty.

## Publish safely

For a new or updated item, prefer versioned asset filenames and keep the catalog ID stable:

1. Build a complete staging copy of the library and run `validate-library` against that staging root.
2. Copy the new versioned PPTX and preview into the live root first. They are harmless while unregistered.
3. Replace the live `catalog.json` with the validated version **last**.
4. Refresh the task pane and verify metadata, preview, details, and insertion.
5. Retain the prior catalog and old registered assets until rollback is no longer needed.

This order avoids a catalog pointing to files that have not arrived. There is still no multi-file transaction or lock; use a maintenance window for a consequential update.

For removal:

1. Remove or deprecate the item in a validated candidate catalog.
2. Publish the catalog and force/confirm a server refresh.
3. Confirm clients no longer receive the item.
4. Only then archive or delete its old assets.

Deprecation is preferable when users still need to discover that a legacy item should no longer be used.

## Make the change visible

The running API compares `catalog.json` modification time and size before catalog reads. A normal catalog edit is loaded on the next list/item request. The task-pane **Refresh** button issues a fresh list request.

The following command validates the library and proves that a fresh index can be built in a short-lived process:

```powershell
npm run reindex-library -- --path "C:\CorporateSlideLibrary"
```

Despite its name, it does not write an index and does not notify a running server. It is a content/operations check.

For a forced refresh of a running local pilot, set `ENABLE_ADMIN_REINDEX=true`, restart the server, then call:

```powershell
Invoke-RestMethod -Method Post -Uri "http://127.0.0.1:3001/api/admin/reindex"
```

The endpoint is intentionally absent by default and has no authentication. Do not expose it on an untrusted network. A server restart also performs a forced initial catalog load.

PPTX and preview bytes are read on each binary request rather than stored in the catalog cache. Replacing a registered asset can therefore become visible without a metadata reload, which is why versioned filenames and catalog-last publication are recommended.

## Verify after publication

- The validator returns `valid: true` for the live root.
- The list contains the intended ID once and shows the intended category/status.
- Search terms find the item.
- The preview matches the source and does not show a placeholder.
- Details show the correct version, owner, status, and update date.
- A compatible PowerPoint host inserts exactly one editable slide.
- The resulting slide has the expected formatting and no sensitive hidden content.
- Server logs show no catalog, unsafe-path, or missing-asset error.

## Roll back

1. Restore the previous `catalog.json` while keeping both old and new assets present.
2. Run `validate-library` against the restored library.
3. Force/confirm server refresh and verify the task pane.
4. Remove failed new assets only after the restored catalog is active.

The product has no built-in version store. Backups, source control for metadata, SharePoint versioning, or another governed repository must provide retention during the MVP.

## Common validation failures

| Symptom | Likely cause | Action |
| --- | --- | --- |
| `Catalog contains invalid JSON` | Trailing comma, quote, or merge error | Fix JSON and rerun validation. |
| Duplicate ID | Two items share the same kebab ID | Choose one stable unique ID and update references. |
| Unsupported extension | Source/preview suffix is outside the allowed set | Export the required format; do not rename incompatible bytes. |
| Unsafe path | Absolute path, backslash, drive, URI, dot/traversal, empty, or control segment | Use a relative forward-slash path inside the root. |
| Asset unavailable | File missing, directory used as a file, permission issue, or broken link | Restore a readable regular file before publishing the catalog. |
| Resolves outside root | Symlink/junction points outside the configured library | Place the real asset inside the governed root. |
| Insert adds multiple slides | Source PPTX violates the one-slide invariant | Recreate and manually verify a one-slide source. |
