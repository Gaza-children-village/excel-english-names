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
- When the service cannot be reached, `TO_ENGLISH` returns `#N/A` and sends no requests for 30 seconds, so recalculation does not slow the PC down.
- A name without a spelling returns `#N/A`; a range of several cells returns `#VALUE!`.
- Names are sent only to the local GCV application. Logs contain counts and failures, never names.

## Name service (planned)

The endpoint does not exist in the GCV application yet. Proposed contract:

```
GET http://127.0.0.1:5555/api/transliterate?name=<UTF-8, URL-encoded Arabic name>
200 {"english": "Mohammed Ahmed Khaled Abu Shamala", "needsReview": true}
```

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

```
python make_addin_release.py "GCV English Names.xlam" --changelog CHANGELOG.md --repo Gaza-children-village/excel-english-names
```

The release tool (in [excel-updater/release-tools](https://github.com/Gaza-children-village/excel-updater/tree/main/release-tools)) checks the file (add-in type, ribbon, no test modules, version listed in `CHANGELOG.md`), writes `manifest.json` with the SHA-256 and size, and prints the `gh release create english-names-v<version> ...` command. Tags are `english-names-v<version>`. Drafts and pre-releases are ignored by the updater.

The file name must stay `GCV English Names.xlam`: workbooks refer to `TO_ENGLISH` through it, and the loader and updater use it.

## Files

| File | Purpose |
|---|---|
| `src/EnglishNames.bas` | `TO_ENGLISH`, cache, pause, ribbon actions |
| `src/HttpNameApi.cls`, `src/INameApi.cls`, `src/NameResult.cls` | Name service (WinHTTP) behind an interface |
| `src/UpdJson.bas` | JSON reader |
| `src/GcvLogClient.bas`, `src/GcvAddin.bas` | Logging through GCV Logger; identity of the add-in |
| `src/TestEnglishNames.bas`, `src/FakeNameApi.cls` | Tests (development only) |

## Part of GCV Tools

Excel add-ins for the schools of Gaza Children's Village. They share one ribbon tab, **Gaza Children Village Tools**, and are installed and updated together.

| Add-in | Repository | What it does |
|---|---|---|
| GCV Logger | [excel-logger](https://github.com/Gaza-children-village/excel-logger) | Logs of all GCV add-ins; show logs, report an issue |
| GCV Updater | [excel-updater](https://github.com/Gaza-children-village/excel-updater) | Installs updates and new tools; home of the catalog, loader and installer |
| GCV Student Lookup | [excel-student-lookup](https://github.com/Gaza-children-village/excel-student-lookup) | Fills student rows from the Civil Registry, on demand |
| GCV English Names | [excel-english-names](https://github.com/Gaza-children-village/excel-english-names) | `=TO_ENGLISH(cell)`: English spelling of Arabic names |

See [CHANGELOG.md](CHANGELOG.md).
