# Demo guide: 2–4 minutes

## Preflight

Do this before the audience arrives, not during the timed demo:

```powershell
npm install
npm run check
npm run dev
```

Keep `npm run dev` running. In a second terminal:

```powershell
npm run sideload
```

Then verify once that:

- PowerPoint opens and the **Open Slide Library** ribbon command is available;
- the task pane loads 10 approved demo items by default;
- `http://127.0.0.1:3001/api/health` returns `{ "status": "ok" }`;
- one test insertion succeeds in the installed PowerPoint build;
- the task-pane console and network panel show no unexpected error.

Open a fresh, noncritical presentation for the actual demo. Because the MVP does not pass a `targetSlideId`, inserted content appears at the beginning of the presentation.

## Live PowerPoint script

### 0:00–0:25 — State the value

Open PowerPoint, press **Home → Open Slide Library**, and say:

> “Instead of finding an old deck and copying from it, the user gets governed, reusable slides inside PowerPoint, with status, version, owner, and freshness visible before insertion.”

Point out that the initial view is filtered to **Approved** content.

### 0:25–1:05 — Find governed content

1. Type `revenue` in search.
2. Select category **Finance**.
3. Point to the one-result count and **Approved** badge.

Say that the search runs against catalog metadata, while the PPTX itself is not downloaded until insertion.

### 1:05–1:45 — Inspect before reuse

Open **Revenue Overview** by clicking its preview or title. Show:

- the enlarged preview;
- version `3.1`;
- update date;
- owner and department;
- tags and approval status.

Close the dialog, or leave it open and use its **Insert slide** button.

### 1:45–2:20 — Insert the real slide

Press **Insert** once. While the button is disabled/spinning, say:

> “Only now does the add-in download the registered one-slide PPTX, convert it for Office.js, and ask PowerPoint to insert native editable content with source formatting.”

Show the success toast, then show the inserted slide at the beginning of the deck. Select a text box or chart element briefly to demonstrate that it is native PowerPoint content, not a screenshot.

### 2:20–3:00 — Show governance and freshness

Return to the task pane, reset filters, and briefly switch status to **Draft** or **Deprecated**. Explain that status is visible rather than hidden in filenames.

Finish with:

> “A content owner replaces the registered one-slide file and preview, updates version/status/date in the catalog, validates it, and users see the change after refresh. The next pilot step is the same API over governed SharePoint storage and enterprise identity.”

Stop here unless the audience asks for architecture or the content workflow.

## Optional 30-second technical appendix

Show, but do not edit live:

- `data/catalog.json` for human-reviewable metadata;
- `data/slides/` and `data/previews/` for the one-item/one-slide storage rule;
- `npm run validate-library` as the content gate;
- [ARCHITECTURE.md](ARCHITECTURE.md) for the storage-adapter and Office boundaries.

Avoid opening a Base64 payload, logging PPTX contents, or changing the demo catalog during the timed flow.

## Browser fallback

If PowerPoint, sideload policy, or the Office host is unavailable, stop the HTTPS processes and run:

```powershell
npm run dev:browser
```

Open <http://localhost:3000> and repeat the catalog/search/filter/details flow. Explicitly point out **Catalog preview mode**. Pressing **Insert** produces an informational message that PowerPoint is required; do not describe that as a successful insertion.

Use this fallback to demonstrate the complete catalog/API UX and then show the insertion sequence in [ARCHITECTURE.md](ARCHITECTURE.md). Be transparent that the external live-host acceptance item remains unchecked.

## Fast recovery

- **No catalog:** confirm `SLIDE_LIBRARY_PATH` is empty/unset for the bundled demo, then restart `npm run dev`.
- **API error:** open `http://127.0.0.1:3001/api/health` and inspect the server terminal's structured error.
- **Certificate warning:** rerun `npm run certs:ensure`, trust the development certificate, and restart the add-in.
- **Add-in missing:** rerun `npm run validate-manifest`, then `npm run sideload`; if required, clear the Office add-in cache or use the documented shared-folder sideload path.
- **Insert unavailable:** confirm the task pane is inside PowerPoint and that the installed host supports `PowerPointApi 1.2`.
- **End the session:** run `npm run sideload:stop`, then stop `npm run dev` with `Ctrl+C`.

Official references: [sideloading on Windows](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins), [manifest validation](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest), and [`insertSlidesFromBase64`](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.presentation).
