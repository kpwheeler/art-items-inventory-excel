# Collection Inventory (Excel)

A private inventory of art, books and other collectibles kept in **one Excel workbook on OneDrive**,
with a **web form** for adding, finding and editing items from an iPhone or PC (Google Chrome).
Books can be filled in from their **ISBN** (typed or scanned) using the free Open Library service.
No photos, no AI. Full specification: [docs/SPEC.md](docs/SPEC.md).

```
workbook/build-workbook.ps1        builds the workbook with desktop Excel
workbook/Collection Inventory.xlsx the empty workbook (template) it produces
web/                               the web form (TypeScript + Vite, static site)
.github/workflows/pages.yml        publishes web/ to GitHub Pages
docs/SPEC.md                       the approved proposal and specification
```

## The workbook

| Tab | What it holds |
| --- | --- |
| How-to | Short instructions |
| Search | One search box (B2) plus filters across every category |
| Art, Books, Items, Misc | One Excel Table per category (`tblArt`, `tblBooks`, `tblItems`, `tblMisc`): shared columns first, then the category's own |
| ISBN Lookup | Book details fetched by Power Query for the ISBNs on the Books tab |
| Valuations | Optional value history by item number |
| Lists | `tblCategories` (name, tab, table, prefix, next item number), `tblDropdowns` and the dropdown lists |
| _Template | Shared columns only; copy it to add a category |

Every column is optional. Type rows straight into a category table, or use the web form; both write
to the same tables.

## Setup

### 1. Put the workbook in OneDrive

Use the committed `workbook/Collection Inventory.xlsx`, or rebuild it (Windows, Microsoft 365 Excel):

```powershell
& .\workbook\build-workbook.ps1
```

Upload it to OneDrive. The default location the form expects is `Documents/Collection Inventory.xlsx`.

**First ISBN refresh (Excel desktop):** Data › Refresh All. Excel asks once for privacy levels
because the query combines the workbook with a web source: choose **Public** for openlibrary.org
(or File › Options › Trust Center › Privacy: ignore privacy levels for this workbook).

### 2. Register the web form with Microsoft (one time, free)

1. Sign in at <https://entra.microsoft.com> (or the Azure portal) with the Microsoft account that
   owns the OneDrive. If it asks you to create a directory or free Azure account first, do that.
2. **App registrations › New registration**
   - Name: `Collection Inventory`
   - Supported account types: **Personal Microsoft accounts only**
   - Redirect URI: platform **Single-page application**,
     `https://kpwheeler.github.io/art-items-inventory-excel/`
3. Under **Authentication**, add a second SPA redirect URI for local testing:
   `http://localhost:5173/art-items-inventory-excel/`
4. Under **API permissions**, add Microsoft Graph › Delegated › `Files.ReadWrite`.
5. Copy the **Application (client) ID**.

### 3. Configure and publish the form

In this GitHub repository: Settings › Secrets and variables › Actions › **Variables**:

| Variable | Value |
| --- | --- |
| `CLIENT_ID` | the Application (client) ID from step 2 |
| `WORKBOOK_PATH` | optional; default `/Documents/Collection Inventory.xlsx` |

Then Actions › *Deploy web form to GitHub Pages* › Run workflow. The form is published at
<https://kpwheeler.github.io/art-items-inventory-excel/>. On iPhone, open it in Chrome and use
Share › Add to Home Screen.

### Local development

```bash
cd web
cp public/config.example.json public/config.json   # then fill in clientId
npm install
npm run dev        # http://localhost:5173/art-items-inventory-excel/
npm run build
```

## Using the form

- **Add item:** pick a category, fill any fields, Save (the item number is assigned) or
  Save & add another (keeps category and acquisition details).
- **Books:** type the ISBN and tap *Look up*, or *Scan* the barcode with the camera. Only empty
  fields are filled.
- **Search:** matches item number, name, creator, description, source, location, tags and notes.
  Tap a result to edit it.
- **Move to category:** on an item's page, e.g. to sort something out of Misc. It gets a new
  number; Notes records the old one.
- **Tidy workbook:** numbers rows typed in Excel without an Item No and fills blank book details
  from their ISBN.
- If a row was changed in Excel after you opened it, the form asks before overwriting.

## Adding a category

See the workbook's How-to tab: copy `_Template`, rename the tab and table, add a row to
`tblCategories`, and add the table to the Search formula in `Search!A11`. The form picks it up
automatically.

## Backups and privacy

- OneDrive keeps version history; also save a copy offline now and then.
- The workbook stays private in your OneDrive. This repository holds only code and the empty
  template; the form contains no secrets and can only reach the workbook after you sign in.

## Status

Phase 1 (workbook) is built and tested in desktop Excel. The web form builds and its ISBN lookup
works; reading and writing the workbook through Microsoft Graph still has to be tested with the
real app registration and a **personal** OneDrive account (Phase 0 check, see the issues).
