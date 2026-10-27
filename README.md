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

Put it in OneDrive at **Inventory/Collection Inventory.xlsx** (a folder named `Inventory` at the top
level of OneDrive). That is where the form opens it; keep the file name. The same folder is a good
place for reference copies of this guide and the specification.

**First ISBN refresh (Excel desktop):** Data › Refresh All. Excel asks once for privacy levels
because the query combines the workbook with a web source: choose **Public** for openlibrary.org
(or File › Options › Trust Center › Privacy: ignore privacy levels for this workbook).

### 2. Register the web form with Microsoft (one time, free)

A personal Microsoft account (Outlook.com, Microsoft 365 Personal/Family) has no Microsoft Entra
directory of its own, and app registration needs one. Signing in to entra.microsoft.com with only a
personal account fails or shows App registrations empty. A free Azure account creates the directory.

**a. Create the directory (free Azure account)**

1. Go to <https://azure.microsoft.com/free> and sign in with the Microsoft account that owns the
   OneDrive holding the workbook.
2. Complete the sign-up. Microsoft asks for a phone number and a card to verify identity; app
   registration itself costs nothing and nothing needs to run in Azure.
3. Sign out, then sign in at <https://entra.microsoft.com> with the same account. A
   **Default Directory** now exists with you as administrator.

**b. Register the app**

1. **Applications › App registrations › New registration**
   - Name: `Collection Inventory`
   - Supported account types: **Personal Microsoft accounts only**. If that option isn't offered,
     choose **Accounts in any organizational directory and personal Microsoft accounts**; the form
     works with either.
   - Redirect URI: platform **Single-page application (SPA)**,
     `https://kpwheeler.github.io/art-items-inventory-excel/`
2. **Authentication** › Single-page application › add a second redirect URI for local testing:
   `http://localhost:5173/art-items-inventory-excel/`
3. **API permissions** › Add a permission › Microsoft Graph › Delegated › `Files.ReadWrite`.
   (`User.Read` is already there; keep it.)
4. **Overview** › copy the **Application (client) ID**. It is not a secret.

The first time you sign in to the form, Microsoft asks you to allow it to read and write your files;
accept once.

If you'd rather not create an Azure account, the alternative is the spec's fallback without the
custom form (Microsoft Forms or Excel-only entry); see docs/SPEC.md, Roadmap › Phase 0.

### 3. Configure and publish the form

In this GitHub repository: Settings › Secrets and variables › Actions › **Variables**:

| Variable | Value |
| --- | --- |
| `CLIENT_ID` | the Application (client) ID from step 2 |
| `WORKBOOK_PATH` | `/Inventory/Collection Inventory.xlsx` (also the default when unset) |

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

- **Add item:** choose the **Category** from the pull-down at the top of the form, fill any
  fields, Save (the item number is assigned) or Save & add another (keeps category and
  acquisition details).
- **Books:** type the ISBN and tap *Look up*, or *Scan* the barcode with the camera. Only empty
  fields are filled.
- **Search:** matches item number, name, creator, description, source, location, tags and notes.
  Tap a result to edit it.
- **Move to another category:** on an item's page, choose a different Category in the pull-down,
  e.g. to sort something out of Misc. It gets a new number; Notes records the old one.
- **Check duplicates:** lists items that may have been entered more than once: same item number,
  ISBN, serial number or name ("likely"), or very similar names ("possible"). Open each to compare,
  and Delete the extra copy (asks first; OneDrive version history can undo).
- **Close:** ends the session and releases the workbook; then close the tab. You stay signed in.
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
