# Collection Inventory App (Excel) — Proposal & Specification

Approved 2026-10-27. Source document: Claude Doc "Collection Inventory App (Excel) — Proposal & Specification".
Implementation notes at the end record where the build differs from the original wording.

## Summary

Keep the whole collection in **one Excel workbook on OneDrive**, with a worksheet tab per category
(Art, Books, Items, Misc). Add items either by typing straight into the workbook or through a
**small web form** that writes into the right tab. The form looks up book details from an ISBN, and
so can the workbook. There are no photos and no AI analysis.

Goals:

- **Two ways in, one source of truth.** Rows added from the web form and rows typed into Excel live in the same tables.
- **Every field optional.** Only the item number is required, and it is assigned automatically.
- **ISBN lookup.** Type or scan an ISBN; title, author, publisher, year and pages fill in from the free Open Library service.
- **One search across every tab** from a Search sheet, plus the normal Excel filters on each tab.
- **New categories without code.** Copy a template tab, rename it, and the web form picks it up.
- **Low cost.** $0 beyond a Microsoft 365 subscription; the form is hosted free.

## How it works

The workbook is the database; the web form and Excel are two windows onto the same tables.

- **Web form** (Chrome on iPhone or PC, Microsoft sign-in) → Microsoft Graph Excel API → workbook.
- **Excel** (desktop or web) → type rows directly on any tab.
- **Open Library** ← ISBN lookups from the form, and from Excel desktop via Power Query (Refresh All).

## Workbook design

`Collection Inventory.xlsx` lives in OneDrive and holds one Excel Table per category tab. Every
category table starts with the same shared columns, in the same order, followed by its own columns.

### Worksheet tabs

| Tab | Excel Table | Purpose |
| --- | --- | --- |
| How-to | — | Short instructions |
| Search | — | One search box and filters across every category |
| Art | `tblArt` | Paintings, prints, sculpture, photographs |
| Books | `tblBooks` | Books, with ISBN lookup |
| Items | `tblItems` | Collectibles and anything else |
| Misc | `tblMisc` | Catch-all for items not sorted yet; move them to another category later |
| ISBN Lookup | `tblIsbn` | Book details fetched from Open Library (refreshed, not typed) |
| Valuations | `tblValuations` | Optional history of value estimates by item number |
| Lists | `tblCategories`, `tblDropdowns`, list tables | Category names, tabs, tables, prefixes and next item number; which columns have dropdowns; the dropdown values |
| _Template | `tblTemplate` | Shared columns only; copied to make a new category |

### Shared columns (every category tab, in this order)

| Column | Type | Notes |
| --- | --- | --- |
| Item No | text | `ART-0001`, `BK-0001`, `ITM-0001`, `MSC-0001`; assigned by the form or Tidy workbook |
| Name | text | Title of the work, book title or item name |
| Creator | text | Artist, author or maker |
| Description | long text | |
| Acquired From Type | dropdown | Artist, Auction, Gallery, Dealer, Retail, Online, Gift, Inheritance, Found, Other |
| Acquired From Name | text | Auction house and lot, gallery, giver |
| Acquired Date | date | |
| Purchase Price | USD | |
| Estimated Value | USD | |
| Value Basis | dropdown | Owner estimate, Appraisal, Auction comparables, Insurance schedule |
| Value Date | date | |
| Location | text | Room, wall, shelf, box |
| Condition | dropdown | Excellent, Good, Fair, Poor |
| Status | dropdown | Owned, Loaned, Sold, Gifted, Donated |
| Tags | text | Comma-separated |
| Notes | long text | Provenance and anything else |
| Entered Via | automatic | Form (blank when typed in Excel) |
| Last Updated | automatic | Set by the form and Tidy workbook |

### Category columns (after the shared columns)

- **Art:** Medium, Dimensions (H x W x D), Units (in/cm, default in), Year Created, Signed, Signature Location, Edition, Framed, Certificate of Authenticity.
- **Books:** ISBN, Author, Publisher, Publication Year, Edition, Printing, Pages, Binding, Signed or Inscribed, Dust Jacket, Lookup Status.
- **Items:** Maker or Brand, Model, Serial Number, Material, Era, Dimensions (H x W x D), Units.
- **Misc:** Suggested Category (where it should go later).

All columns may be blank. Dropdowns use Excel data validation pointing at the Lists tab; the web
form reads the same lists from `tblDropdowns`.

## Web form

1. **Sign in** with the Microsoft account that owns the OneDrive workbook. The page keeps no passwords or keys.
2. **Pick a category.** Categories come from `tblCategories`; the fields shown come from that table's headers.
3. **Books: enter or scan the ISBN**, then Look up. Only empty fields are filled; unknown ISBNs show "Not found". Scanning uses the camera through a JavaScript barcode library (ZXing) and works in Chrome on Windows (webcam) and iPhone; no image is stored.
4. **Fill in what you know and Save.** The form assigns the next item number, sets Entered Via and Last Updated, and adds the row to the right tab.
5. **Find and edit.** Search lists matching items from all tabs; open one to change it and save it back to the same row.

Details: Save & add another keeps category and acquisition fields; Units defaults to inches; target
browser is Google Chrome on Windows and iPhone (Share › Add to Home Screen for an icon); hosted free
on GitHub Pages as a static site; if a row changed in Excel since it was opened, the form asks
before overwriting.

### Moving items between categories

- **Web form:** open the item › Move to category. The row moves to that tab with all shared fields, gets the next number in the new prefix, and Notes records "Moved from MSC-0004 on <date>".
- **Excel:** cut the row, paste it at the bottom of the target table, clear Item No; Tidy workbook assigns the new number.
- **Renaming Misc itself** is a tab rename plus a change to its row in `tblCategories`; existing numbers keep the `MSC` prefix.

## Spreadsheet entry and ISBN lookup in Excel

Typing into a category tab is fully supported. The workbook stays a plain `.xlsx` with no macros.

- **Item numbers for typed rows:** the Lists tab shows the next free number per category; or leave Item No blank and run Tidy workbook in the form.
- **ISBN lookup for typed rows:**
  - *Tidy workbook* (web form, any device): fills blank book fields from Open Library and sets Lookup Status.
  - *Refresh All* (Excel desktop): Power Query fetches details for every ISBN on the Books tab into the ISBN Lookup tab. The first refresh asks for privacy levels once.

Open Library is free and needs no account; coverage is good but not complete.

## Search

The Search sheet searches every category at once: a word in B2 matches Item No, Name, Creator,
Description, Acquired From Name, Location, Tags or Notes. Optional filters: Category, Status,
Location, minimum and maximum Estimated Value. Results list the shared columns plus Category, sorted
by name, with a match count and total estimated value. One formula (`LET`, `VSTACK`, `FILTER`,
`SEARCH`) needs Microsoft 365 Excel. The web form offers the same search on the phone.

## Adding categories, backups and security

1. Copy the `_Template` tab and rename it (e.g. Coins).
2. Rename its table (e.g. `tblCoins`) and add the category's own columns after the shared ones.
3. Add a row to `tblCategories` (name, tab, table, prefix).
4. Add the table to the Search formula in `Search!A11`.

OneDrive keeps version history; save an offline copy now and then. The workbook stays private in
OneDrive; the form reaches it only after sign-in (Graph scope `Files.ReadWrite`) and holds no secrets.

## Costs

| Component | Cost |
| --- | --- |
| Microsoft 365 (already owned) | $0 extra |
| GitHub Pages hosting (public repository) | $0 |
| Microsoft app registration | $0 |
| Open Library | $0 |

## Decisions

| Question | Decision |
| --- | --- |
| Microsoft account type | Personal |
| Microsoft 365 | Already owned |
| Existing spreadsheet to import | None |
| Categories at launch | Art, Books, Items, Misc |
| Who enters items | Just the owner |
| Barcode scanning | Yes, Phase 2 |
| Currency and units | USD, inches |
| GitHub repository | `kpwheeler/art-items-inventory-excel` |

## Roadmap

1. **Phase 0 — Check.** Confirm the form can write to an Excel table in a personal OneDrive. Fallback: Microsoft Forms into one sheet.
2. **Phase 1 — Workbook.** Tabs, tables, lists and dropdowns, next-number formulas, Search sheet, Power Query ISBN lookup, `_Template`, How-to.
3. **Phase 2 — Web form.** Sign-in, dynamic categories and fields, ISBN lookup and scan, save with numbering, search and edit, move to category.
4. **Phase 3 — Polish.** Tidy workbook, change warnings, Google Books as a second ISBN source, printable inventory report.

## Implementation notes (differences from the original wording)

- **Repository is public**, not private: GitHub Pages cannot publish from a private repository on a
  free GitHub plan. The repository holds only code and the empty template; no personal data.
- **Column names avoid `/`** (Name instead of "Name / Title", "Maker or Brand", "Signed or
  Inscribed") because `/` needs escaping in Excel structured references.
- **No "Found Title / Found Author" columns on the Books tab.** Formula columns inside the table
  would break rows added by the form. Looked-up details appear on the ISBN Lookup tab, and Tidy
  workbook copies them into blank Books cells.
- **Lists tab adds `tblDropdowns`** (column → list) so the form offers exactly the sheet's dropdowns.
- **Tidy workbook and change warnings** (Phase 3 items) were built early, with the form.
