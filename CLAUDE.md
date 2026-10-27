# Collection Inventory (Excel) — project conventions

- The workbook (`workbook/Collection Inventory.xlsx`) is the database. It stays a plain `.xlsx`:
  no macros, because the web form writes to it through the Microsoft Graph Excel API, which
  does not support `.xlsm`.
- Regenerate the workbook with `workbook/build-workbook.ps1` (desktop Excel via COM). Change the
  script, not the binary by hand, so the two never drift apart.
- Every category table starts with the same shared columns, in this order: Item No, Name, Creator,
  Description, Acquired From Type, Acquired From Name, Acquired Date, Purchase Price, Estimated
  Value, Value Basis, Value Date, Location, Condition, Status, Tags, Notes, Entered Via, Last
  Updated. The Search formula and the web form depend on this order and these names. Avoid `/` and
  other special characters in column names (they need escaping in structured references).
- Every item field is optional; only Item No is filled automatically (`PREFIX-0001`).
- Lookups (ISBN) and tidy-ups only fill blank cells. They never overwrite what the user typed.
- Categories, dropdown lists and their columns come from the workbook (`tblCategories`,
  `tblDropdowns`); the form must not hard-code them.
- No secrets in the repo. The Microsoft app (client) ID and workbook path are configuration, not
  secrets: `web/public/config.json` locally (git-ignored), repository variables in CI.
- Target browser: Google Chrome on Windows and iPhone. Barcode scanning uses ZXing because Chrome
  on those platforms has no BarcodeDetector.
- Web form: TypeScript strict, no UI framework. `npm run build` (in `web/`) must pass before pushing.
