# Changelog - gcv/english-names (GCV English Names add-in)

## 0.2.0

- Ready for the name service of the GCV desktop app (after v2.0.0; until then the app answers 503 and the add-in pauses as below): requests send `X-GCV-Excel-Addin` and, when the app has written one, its key `X-GCV-Api-Key` (from `HKCU\Software\GCV Tools\LocalApiKey`; never logged).
- 401, 503 and other HTTP errors now start the same 30-second pause as a lost connection (before, every cell asked again).
- "Try again" first explains the last problem (key refused: open or update the app; service not ready; app not running), using `GET /api/health` for the app version.
- The endpoint is confirmed by the app's contract; no longer "planned".

## 0.1.0

- Worksheet function =TO_ENGLISH(cell): English spelling (passport style) of an Arabic name from the GCV application's name service - a suggestion to review.
- Ribbon: "freeze reviewed names" turns selected TO_ENGLISH formulas into fixed text; "try again" asks the service again.
- Light on the PC: each name is asked once per session; when the service is down, no requests for 30 seconds (#N/A).
- Needs the endpoint GET /api/transliterate?name=... in the GCV application (not built yet).
