# Collection Inventory App (Excel) — Proposal & Specification

Approved 2026-10-27; updated 2026-10-27 after the first build and test. Source document: Claude Doc
"Collection Inventory App (Excel) — Proposal & Specification". Implementation notes at the end record
where the build differs from the original wording.

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

`Collection Inventory.xlsx` lives in OneDrive at **Apps › Inventory › Collection Inventory.xlsx**, the
form's own app folder and the only folder it can reach (see Security below). The top-level OneDrive
folder named Inventory keeps reference copies of this specification and the setup guide. The
workbook holds one Excel Table per category tab. Every
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
2. **Choose the Category** from the pull-down that is the first field of the form. Categories come from `tblCategories`; the fields shown come from that table's headers. Switching category while adding keeps what was typed in the shared fields.
3. **Books: enter or scan the ISBN**, then Look up. Only empty fields are filled; unknown ISBNs show "Not found". Scanning uses the camera through a JavaScript barcode library (ZXing) and works in Chrome on Windows (webcam) and iPhone; no image is stored.
4. **Fill in what you know and Save.** The form assigns the next item number, sets Entered Via and Last Updated, and adds the row to the right tab.
5. **Find and edit.** Search lists matching items from all tabs; open one to change it and save it back to the same row.

6. **Check duplicates** (menu) lists items that may have been entered more than once (see below).
7. **Close** (menu, right end) ends the session: it releases the workbook in Excel and shows "Finished" with an Open again button. Chrome usually doesn't let a page close its own tab, so the tab is closed by hand. It does not sign out.

Details: Save & add another keeps category and acquisition fields; Units defaults to inches; target
browser is Google Chrome on Windows and iPhone (Share › Add to Home Screen for an icon); hosted free
on GitHub Pages as a static site; if a row changed in Excel since it was opened, the form asks
before overwriting.

Reliability: the form opens one Excel session and reuses it for every request (reopening it when it
expires), because opening the workbook afresh on each request caused "workbook cannot be opened"
timeouts (HTTP 504). Timeouts and throttling (429/502/503/504) are retried automatically for requests
that are safe to repeat. Adding a row is never retried blindly: after a failed add the form checks
whether the row arrived. While saving, the buttons show "Saving…" and further clicks are ignored.

### Check duplicates

Re-reads every category tab and groups items that may be the same thing entered twice, across
categories:

| Group | Matched on |
| --- | --- |
| Likely | Same Item No; same ISBN (hyphens and spaces ignored); same Serial Number; same Name ignoring capitals, punctuation and a leading "The", "A" or "An" |
| Possible | Very similar names (letter-pair similarity of 85% or more), e.g. "Paper Koi Sculpture" and "Paper Koi Sculptures" |

Each item shows its number, category, creator, location, acquired date and Entered Via, with **Open**
(compare and edit) and **Delete** (after a confirmation; OneDrive version history can undo; the last
row of a table is cleared instead of deleted, because a table keeps one row).

### Moving items between categories

- **Web form:** on an item's page, choose a different Category in the pull-down and confirm. The row moves to that tab with all shared fields, gets the next number in the new prefix, and Notes records "Moved from MSC-0004 on <date>". Unsaved edits on the page are not carried over.
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
OneDrive; the form reaches it only after sign-in and holds no secrets.

**Security: least privilege.** The form signs in with Microsoft Graph `Files.ReadWrite.AppFolder`,
which allows reading and writing only the app's own folder, OneDrive › Apps › Inventory (named after
the app registration "Inventory"). The rest of the OneDrive, including tax, banking, medical and
legal folders, is out of reach, so a stolen token cannot expose it. Personal Microsoft accounts have
no permission for a folder of the owner's choosing; the app folder is the narrowest available. Any
earlier consent to `Files.ReadWrite` is removed at account.live.com/consent/Manage. Microsoft
documents the Excel workbook API with `Files.ReadWrite`; if it refuses the app-folder permission,
the repository variable `ACCESS=allFiles` restores full access without a code change.

## Costs

| Component | Cost |
| --- | --- |
| Microsoft 365 (already owned) | $0 extra |
| GitHub Pages hosting (public repository) | $0 |
| Microsoft app registration (needs a free Azure account, see Decisions) | $0 |
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
| GitHub repository | `kpwheeler/art-items-inventory-excel` (public) |
| Workbook location | OneDrive › Apps › Inventory › Collection Inventory.xlsx (the form's app folder); reference documents stay in OneDrive › Inventory |
| Form's access to OneDrive | Only its app folder (`Files.ReadWrite.AppFolder`), not all files |
| App registration | A personal Microsoft account has no Microsoft Entra directory, so a free Azure account was created to get one. App "Inventory", personal accounts, SPA redirect `https://kpwheeler.github.io/art-items-inventory-excel/`, Graph `Files.ReadWrite.AppFolder` |
| Target browser | Google Chrome on Windows and iPhone |

## Roadmap

1. **Phase 0 — Check.** Confirm the form can write to an Excel table in a personal OneDrive. Fallback: Microsoft Forms into one sheet. *Status: sign-in and opening the workbook work; the first save timed out, which led to the session-and-retry change; saving to be re-tested.*
2. **Phase 1 — Workbook.** Tabs, tables, lists and dropdowns, next-number formulas, Search sheet, Power Query ISBN lookup, `_Template`, How-to. *Status: built and tested in desktop Excel.*
3. **Phase 2 — Web form.** Sign-in, dynamic categories and fields, Category pull-down, ISBN lookup and scan, save with numbering, search and edit, move by changing category, Close. *Status: built and deployed; being tested.*
4. **Phase 3 — Polish.** Tidy workbook, change warnings, check duplicates, Google Books as a second ISBN source, printable inventory report. *Status: the first three built early; the last two not started.*

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
- **Workbook moved to OneDrive › Inventory** instead of Documents, at the owner's request, then to
  **OneDrive › Apps › Inventory** so the form could use the app-folder permission instead of access to
  all files (security review, 2026-10-27).
- **Category is a pull-down at the top of the form**, and moving an item is done by changing it,
  replacing the separate "Move to category" control.
- **Added after the spec:** the Close button, the Check duplicates page, and the single Excel session
  with automatic retries (fix for HTTP 504 "workbook cannot be opened" on save).
