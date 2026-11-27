# GCV English Names

![Release](https://img.shields.io/github/v/release/Gaza-children-village/excel-english-names?filter=english-names-v*&label=release)
![Excel](https://img.shields.io/badge/Excel-2016%2B%20Windows-217346)
![Status](https://img.shields.io/badge/status-preview-orange)

Excel add-in (`GCV English Names.xlam`) with one worksheet function:

```
=TO_ENGLISH(C4)      ->   Mohammed Ahmed Khaled Abu Shamala
```

English spelling (passport style, four parts, Title Case) of the Arabic name in a cell, from the name service of the GCV application. The result is a **suggestion to review**.

## Ribbon

| Button | What it does |
|---|---|
| **تثبيت الأسماء المراجعة** (freeze reviewed names) | Turns the selected `TO_ENGLISH` formulas into fixed text, so a reviewed spelling never changes and the file works without the add-in |
| **إعادة المحاولة** (try again) | Asks the service again for the selected cells (after `#N/A`) |

## Behaviour

- Each name is asked once per Excel session (cache).
- When the service cannot answer (app not running, key refused `401`, not ready `503`, another HTTP error), `TO_ENGLISH` returns `#N/A` and sends no requests for 30 seconds, so recalculation does not slow the PC down. A worksheet function shows no messages: **إعادة المحاولة** explains the last problem, then asks again.
- A name without a spelling returns `#N/A`; a range of several cells returns `#VALUE!`.
- Names are sent only to the local GCV application. Logs contain counts and failures, never names.

## Name service (GCV desktop app v1.10.0 and later)

Contract: `docs/local-api.md` in the GCV desktop app repository.

```
GET http://127.0.0.1:5555/api/transliterate?name=<UTF-8, URL-encoded Arabic name>
200 {"english": "Mohammed Ahmed Khaled Abu Shamala", "needsReview": true}   ("english" empty: no spelling)
401 key refused · 503 not ready · GET /api/health after a failure
```

Requests carry `X-GCV-Excel-Addin: gcv/english-names/<version>` and, once the app has written one, its key `X-GCV-Api-Key` (read from `HKCU\Software\GCV Tools\LocalApiKey`, never logged). The app requires the key only after all PCs have 0.2.0 or later.

The base address can be changed per PC in the registry: `HKCU\Software\VB and VBA Program Settings\GCV Tools\EnglishNames\ApiBaseUrl`.

## Installing

Secretaries do not install add-ins one by one. The **GCV Tools installer** (see [excel-updater](https://github.com/Gaza-children-village/excel-updater#installing-on-a-pc)) installs all of them for the current Windows user, and the **update tools** button keeps them current.

## Building from source

Excel cannot load a VBA project created outside Excel, so the add-in is assembled in Excel from the files in this repository:

1. Copy `GCV English Names - shell.xlam` and rename the copy to `GCV English Names.xlam` (the shell already contains the add-in file type and the ribbon).
2. Open it, press **Alt+F11** and select its VBAProject.
3. **File › Import File**: every file in `src/`.
4. Double-click **ThisWorkbook** and paste the contents of `ThisWorkbook.code`.
5. **Debug › Compile VBAProject**, then save with **Ctrl+S**. Saving records the add-in's id, version and file name in the file.

Source files are ASCII with CRLF line endings. Arabic text is stored as Unicode code points (`FromCodePoints("0627 ...")`), because the VBA editor cannot keep Arabic characters.

## Tests

Run `RunEnglishNamesTests` from the VBA editor (click inside the procedure and press **F5**). Results appear in a message box and in the Immediate window (**Ctrl+G**). The tests use fakes and temporary files: no network, no real data.

Before a release, remove the development modules (`TestEnglishNames`, `FakeNameApi`), compile and save.

## Releasing

Releases are checked and published by GitHub Actions ([`.github/workflows/ci.yml`](.github/workflows/ci.yml)); only the build happens by hand, because it needs Excel.

1. Raise `ADDIN_VERSION` in `src/GcvAddin.bas` and add a `## <version>` section to `CHANGELOG.md`.
2. Build the add-in in Excel (see above) **without** the test modules, and save it.
3. Upload it as a **draft** release:
   ```
   gh release create english-names-v<version> "GCV English Names.xlam" --draft --repo Gaza-children-village/excel-english-names
   ```
4. Merge to `main` and push the tag `english-names-v<version>` (through Gitea; the mirror carries it to GitHub).
5. The workflow checks the uploaded file (same id and version as the tag, add-in type, ribbon, no test modules, CHANGELOG section), adds `manifest.json` (SHA-256 and size), renames the file to its asset name, sets the release notes from `CHANGELOG.md` and **publishes**. If anything is wrong, the release stays a draft (the updater ignores drafts) and the job shows why.

Uploaded the draft after pushing the tag? Re-run the failed **release** job in the Actions tab, or `gh workflow run ci.yml -f tag=english-names-v<version> --repo Gaza-children-village/excel-english-names`.

Every push also runs the source checks: VBA files ASCII with CRLF, module names equal to file names, `Option Explicit`, the shared ribbon tab, a valid identity in `GcvAddin.bas` with a CHANGELOG section. Run them locally with `python .github/scripts/gcv_ci.py check`.

## Files

| File | Purpose |
|---|---|
| `src/EnglishNames.bas` | `TO_ENGLISH`, cache, pause, ribbon actions |
| `src/HttpNameApi.cls`, `src/INameApi.cls`, `src/NameResult.cls` | Name service (WinHTTP) behind an interface |
| `src/UpdJson.bas` | JSON reader |
| `src/GcvLocalApi.bas` | Headers (add-in, key) and health check for the GCV app; same file in Student Lookup |
| `src/GcvLogClient.bas`, `src/GcvAddin.bas` | Logging through GCV Logger; identity of the add-in |
| `src/TestEnglishNames.bas`, `src/FakeNameApi.cls`, `src/TestGcvLocalApi.bas`, `src/FakeHeaderSink.cls` | Tests (development only) |

## Part of GCV Tools

Excel add-ins for the schools of Gaza Children's Village. They share one ribbon tab, **Gaza Children Village Tools**, and are installed and updated together.

| Add-in | Repository | What it does |
|---|---|---|
| GCV Logger | [excel-logger](https://github.com/Gaza-children-village/excel-logger) | Logs of all GCV add-ins; show logs, report an issue |
| GCV Updater | [excel-updater](https://github.com/Gaza-children-village/excel-updater) | Installs updates and new tools; home of the catalog and loader |
| GCV Student Lookup | [excel-student-lookup](https://github.com/Gaza-children-village/excel-student-lookup) | Fills student rows from the Civil Registry, on demand |
| GCV English Names | [excel-english-names](https://github.com/Gaza-children-village/excel-english-names) | `=TO_ENGLISH(cell)`: English spelling of Arabic names |

See [CHANGELOG.md](CHANGELOG.md).
