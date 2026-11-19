# Changelog - gcv/english-names (GCV English Names add-in)

## 0.1.0

- Worksheet function =TO_ENGLISH(cell): English spelling (passport style) of an Arabic name from the GCV application's name service - a suggestion to review.
- Ribbon: "freeze reviewed names" turns selected TO_ENGLISH formulas into fixed text; "try again" asks the service again.
- Light on the PC: each name is asked once per session; when the service is down, no requests for 30 seconds (#N/A).
- Needs the endpoint GET /api/transliterate?name=... in the GCV application (not built yet).
