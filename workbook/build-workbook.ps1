<#
.SYNOPSIS
  Builds "Collection Inventory.xlsx" from scratch using desktop Excel (COM automation).

.DESCRIPTION
  Creates the category tabs (Art, Books, Items, Misc) as Excel Tables with the shared
  columns first, the Lists tab that drives dropdowns and item numbering, the Search
  sheet, the ISBN Lookup Power Query, Valuations, How-to and _Template tabs.

  Requires Windows with Microsoft 365 Excel (for LET, VSTACK, FILTER and Power Query).

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File workbook\build-workbook.ps1
#>
param(
    [string]$OutFile = (Join-Path $PSScriptRoot 'Collection Inventory.xlsx')
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------- definitions

# Shared columns, in order, at the start of every category table.
$SharedColumns = @(
    'Item No', 'Name', 'Creator', 'Description',
    'Acquired From Type', 'Acquired From Name', 'Acquired Date',
    'Purchase Price', 'Estimated Value', 'Value Basis', 'Value Date',
    'Location', 'Condition', 'Status', 'Tags', 'Notes',
    'Entered Via', 'Last Updated'
)

$Categories = @(
    @{ Name = 'Art';   Table = 'tblArt';   Prefix = 'ART'
       Columns = @('Medium', 'Dimensions (H x W x D)', 'Units', 'Year Created', 'Signed',
                   'Signature Location', 'Edition', 'Framed', 'Certificate of Authenticity') }
    @{ Name = 'Books'; Table = 'tblBooks'; Prefix = 'BK'
       Columns = @('ISBN', 'Author', 'Publisher', 'Publication Year', 'Edition', 'Printing',
                   'Pages', 'Binding', 'Signed or Inscribed', 'Dust Jacket', 'Lookup Status') }
    @{ Name = 'Items'; Table = 'tblItems'; Prefix = 'ITM'
       Columns = @('Maker or Brand', 'Model', 'Serial Number', 'Material', 'Era',
                   'Dimensions (H x W x D)', 'Units') }
    @{ Name = 'Misc';  Table = 'tblMisc';  Prefix = 'MSC'
       Columns = @('Suggested Category') }
)

# Single-column lists on the Lists tab: table name -> header, values.
$Lists = [ordered]@{
    tblAcquiredFrom  = @{ Header = 'Acquired From Type'; Values = @('Artist', 'Auction', 'Gallery', 'Dealer', 'Retail', 'Online', 'Gift', 'Inheritance', 'Found', 'Other') }
    tblValueBasis    = @{ Header = 'Value Basis';        Values = @('Owner estimate', 'Appraisal', 'Auction comparables', 'Insurance schedule') }
    tblCondition     = @{ Header = 'Condition';          Values = @('Excellent', 'Good', 'Fair', 'Poor') }
    tblStatus        = @{ Header = 'Status';             Values = @('Owned', 'Loaned', 'Sold', 'Gifted', 'Donated') }
    tblUnits         = @{ Header = 'Units';              Values = @('in', 'cm') }
    tblYesNo         = @{ Header = 'Yes or No';          Values = @('Yes', 'No') }
    tblBinding       = @{ Header = 'Binding';            Values = @('Hardcover', 'Paperback', 'Leather', 'Spiral', 'Other') }
    tblLookupStatus  = @{ Header = 'Lookup Status';      Values = @('Found', 'Not found', 'Edited') }
}

# Which columns get a dropdown, and the list (structured reference) it uses.
# The web form reads this same table, so the sheet and the form offer the same choices.
$Dropdowns = [ordered]@{
    'Acquired From Type'          = 'tblAcquiredFrom'
    'Value Basis'                 = 'tblValueBasis'
    'Condition'                   = 'tblCondition'
    'Status'                      = 'tblStatus'
    'Units'                       = 'tblUnits'
    'Signed'                      = 'tblYesNo'
    'Framed'                      = 'tblYesNo'
    'Certificate of Authenticity' = 'tblYesNo'
    'Dust Jacket'                 = 'tblYesNo'
    'Binding'                     = 'tblBinding'
    'Lookup Status'               = 'tblLookupStatus'
    'Suggested Category'          = 'tblCategories[Name]'
}

$CurrencyColumns = @('Purchase Price', 'Estimated Value')
$DateColumns     = @('Acquired Date', 'Value Date')
$TextColumns     = @('Item No', 'ISBN', 'Serial Number', 'Publication Year', 'Year Created')
$WideColumns     = @('Name', 'Description', 'Notes', 'Acquired From Name')

$IsbnQuery = @'
let
    Source = Excel.CurrentWorkbook(){[Name = "tblBooks"]}[Content],
    Cleaned = List.Transform(Source[ISBN], each Text.Select(Text.From(_ ?? ""), {"0".."9", "X", "x"})),
    Isbns = List.Distinct(List.Select(Cleaned, each Text.Length(_) = 10 or Text.Length(_) = 13)),
    Names = (items) => if items = null then null else Text.Combine(List.Transform(items, each Record.FieldOrDefault(_, "name", "")), ", "),
    Fetch = (isbn as text) =>
        let
            Raw = try Json.Document(Web.Contents("https://openlibrary.org/api/books",
                      [Query = [bibkeys = "ISBN:" & isbn, format = "json", jscmd = "data"]])) otherwise null,
            Book = if Raw = null then null else Record.FieldOrDefault(Raw, "ISBN:" & isbn, null)
        in
            [ISBN = isbn,
             Title = if Book = null then null else Record.FieldOrDefault(Book, "title", null),
             Author = if Book = null then null else Names(Record.FieldOrDefault(Book, "authors", null)),
             Publisher = if Book = null then null else Names(Record.FieldOrDefault(Book, "publishers", null)),
             #"Publication Year" = if Book = null then null else Record.FieldOrDefault(Book, "publish_date", null),
             Pages = if Book = null then null else Record.FieldOrDefault(Book, "number_of_pages", null),
             #"Lookup Status" = if Book = null then "Not found" else "Found"],
    Columns = {"ISBN", "Title", "Author", "Publisher", "Publication Year", "Pages", "Lookup Status"},
    Result = if List.IsEmpty(Isbns)
             then #table(Columns, {})
             else Table.FromRecords(List.Transform(Isbns, Fetch), Columns)
in
    Result
'@

# ---------------------------------------------------------------- helpers

function Set-Headers($ws, [string[]]$headers, [int]$row = 1, [int]$col = 1) {
    for ($i = 0; $i -lt $headers.Count; $i++) { $ws.Cells.Item($row, $col + $i).Value2 = $headers[$i] }
}

function New-Table($ws, [string]$name, [string[]]$headers, [int]$row = 1, [int]$col = 1, [string]$style = 'TableStyleMedium2') {
    Set-Headers $ws $headers $row $col
    $range = $ws.Range($ws.Cells.Item($row, $col), $ws.Cells.Item($row + 1, $col + $headers.Count - 1))
    $lo = $ws.ListObjects.Add(1, $range, $null, 1)   # xlSrcRange, xlYes (has headers)
    $lo.Name = $name
    $lo.TableStyle = $style
    return $lo
}

function Format-CategoryTable($lo) {
    foreach ($lc in $lo.ListColumns) {
        $n = $lc.Name
        $body = $lc.DataBodyRange
        if ($CurrencyColumns -contains $n) { $body.NumberFormat = '$#,##0.00' }
        elseif ($DateColumns -contains $n) { $body.NumberFormat = 'yyyy-mm-dd' }
        elseif ($n -eq 'Last Updated')     { $body.NumberFormat = 'yyyy-mm-dd hh:mm' }
        elseif ($TextColumns -contains $n) { $body.NumberFormat = '@' }

        if ($Dropdowns.Contains($n)) {
            $body.Validation.Delete()
            # xlValidateList = 3, xlValidAlertStop = 1, xlBetween = 1
            $body.Validation.Add(3, 1, 1, "=INDIRECT(""$($Dropdowns[$n])"")")
            $body.Validation.IgnoreBlank = $true
        }
        $lc.Range.ColumnWidth = if ($WideColumns -contains $n) { 36 } else { [Math]::Max(12, $n.Length + 4) }
    }
    $lo.HeaderRowRange.WrapText = $false
}

function Add-Sheet($wb, [string]$name) {
    $last = $wb.Worksheets.Item($wb.Worksheets.Count)
    $ws = $wb.Worksheets.Add([Type]::Missing, $last)
    $ws.Name = $name
    return $ws
}

# ---------------------------------------------------------------- build

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

    $howTo = $wb.Worksheets.Item(1); $howTo.Name = 'How-to'
    $search = Add-Sheet $wb 'Search'
    $catSheets = @{}
    foreach ($c in $Categories) { $catSheets[$c.Name] = Add-Sheet $wb $c.Name }
    $isbnWs = Add-Sheet $wb 'ISBN Lookup'
    $valWs = Add-Sheet $wb 'Valuations'
    $listWs = Add-Sheet $wb 'Lists'
    $tplWs = Add-Sheet $wb '_Template'

    # Lists tab: categories (with next free item number), dropdown map, list tables.
    $catLo = New-Table $listWs 'tblCategories' @('Name', 'Tab', 'Table', 'Prefix', 'Next Item No') 1 1 'TableStyleLight9'
    for ($i = 0; $i -lt $Categories.Count; $i++) {
        $c = $Categories[$i]
        if ($i -gt 0) { [void]$catLo.ListRows.Add() }
        $r = $catLo.ListRows.Item($i + 1).Range
        $r.Cells.Item(1, 1).Value2 = $c.Name
        $r.Cells.Item(1, 2).Value2 = $c.Name
        $r.Cells.Item(1, 3).Value2 = $c.Table
        $r.Cells.Item(1, 4).Value2 = $c.Prefix
    }
    $catLo.ListColumns.Item('Next Item No').DataBodyRange.Formula2 =
        '=LET(p,[@Prefix],ids,INDIRECT([@Table]&"[Item No]"),n,IFERROR(--SUBSTITUTE(ids,p&"-",""),0),p&"-"&TEXT(MAX(n)+1,"0000"))'

    $ddLo = New-Table $listWs 'tblDropdowns' @('Column', 'List') 1 7 'TableStyleLight9'
    $i = 0
    foreach ($k in $Dropdowns.Keys) {
        if ($i -gt 0) { [void]$ddLo.ListRows.Add() }
        $r = $ddLo.ListRows.Item($i + 1).Range
        $r.Cells.Item(1, 1).Value2 = $k
        $r.Cells.Item(1, 2).Value2 = $Dropdowns[$k]
        $i++
    }

    $col = 10
    foreach ($t in $Lists.Keys) {
        $def = $Lists[$t]
        $lo = New-Table $listWs $t @($def.Header) 1 $col 'TableStyleLight9'
        for ($j = 0; $j -lt $def.Values.Count; $j++) {
            if ($j -gt 0) { [void]$lo.ListRows.Add() }
            $lo.ListRows.Item($j + 1).Range.Cells.Item(1, 1).Value2 = $def.Values[$j]
        }
        $col += 2
    }
    $listWs.Columns.Item('A:Z').ColumnWidth = 18

    # Category tabs.
    foreach ($c in $Categories) {
        $lo = New-Table $catSheets[$c.Name] $c.Table ($SharedColumns + $c.Columns)
        Format-CategoryTable $lo
    }

    # Template tab: shared columns only; copy this tab to add a category.
    $tplLo = New-Table $tplWs 'tblTemplate' $SharedColumns
    Format-CategoryTable $tplLo

    # Valuations.
    $valLo = New-Table $valWs 'tblValuations' @('Item No', 'Value Date', 'Estimated Value', 'Value Basis', 'Notes')
    $valLo.ListColumns.Item('Value Date').DataBodyRange.NumberFormat = 'yyyy-mm-dd'
    $valLo.ListColumns.Item('Estimated Value').DataBodyRange.NumberFormat = '$#,##0.00'
    $valLo.ListColumns.Item('Value Basis').DataBodyRange.Validation.Add(3, 1, 1, '=INDIRECT("tblValueBasis")')
    $valWs.Columns.Item('A:E').ColumnWidth = 20

    # ISBN Lookup: Power Query from tblBooks[ISBN] to Open Library, loaded as tblIsbn.
    [void]$wb.Queries.Add('ISBN Lookup', $IsbnQuery)
    $conn = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="ISBN Lookup";Extended Properties=""'
    $qlo = $isbnWs.ListObjects.Add(0, $conn, $null, 1, $isbnWs.Range('A1'))   # xlSrcExternal
    $qlo.QueryTable.CommandType = 2                                            # xlCmdSql
    $qlo.QueryTable.CommandText = 'SELECT * FROM [ISBN Lookup]'
    $qlo.QueryTable.BackgroundQuery = $false
    $qlo.QueryTable.RefreshStyle = 1                                           # xlInsertDeleteCells
    $qlo.QueryTable.PreserveColumnInfo = $true
    [void]$qlo.QueryTable.Refresh($false)
    $qlo.Name = 'tblIsbn'
    $qlo.TableStyle = 'TableStyleLight9'
    $isbnWs.Columns.Item('A:G').ColumnWidth = 22

    # Search sheet: one box + filters over the shared columns of every category table.
    $search.Range('A1').Value2 = 'Search the whole collection'
    $search.Range('A1').Font.Size = 14
    $search.Range('A1').Font.Bold = $true
    $labels = @('Search for', 'Category', 'Status', 'Location contains', 'Min value', 'Max value')
    for ($i = 0; $i -lt $labels.Count; $i++) { $search.Cells.Item(2 + $i, 1).Value2 = $labels[$i] }
    $search.Range('B2:B7').Interior.Color = 0xCCF2FF   # light yellow input cells (BGR)
    $search.Range('B2:B7').Borders.LineStyle = 1
    $search.Range('B3').Validation.Add(3, 1, 1, '=INDIRECT("tblCategories[Name]")')
    $search.Range('B4').Validation.Add(3, 1, 1, '=INDIRECT("tblStatus")')
    $search.Range('B6:B7').NumberFormat = '$#,##0'
    $search.Range('D2').Value2 = 'Matches'
    $search.Range('D3').Value2 = 'Total estimated value'
    $search.Range('E2').Formula2 = '=IF(A11="No matches",0,ROWS(A11#))'
    $search.Range('E3').Formula2 = '=IF(A11="No matches",0,SUM(INDEX(A11#,0,10)))'
    $search.Range('E3').NumberFormat = '$#,##0.00'
    $search.Range('D2:D3').Font.Bold = $true

    Set-Headers $search (@('Category') + $SharedColumns[0..15]) 10 1
    $hdr = $search.Range('A10:Q10')
    $hdr.Font.Bold = $true
    $hdr.Interior.Color = 0x7F3F1F
    $hdr.Font.Color = 0xFFFFFF

    # One VSTACK line per category. When adding a category, add a matching line here (see How-to).
    $search.Range('A11').Formula2 = @'
=LET(
 d, VSTACK(tblArt[[Item No]:[Notes]], tblBooks[[Item No]:[Notes]], tblItems[[Item No]:[Notes]], tblMisc[[Item No]:[Notes]]),
 cat, VSTACK(EXPAND("Art",ROWS(tblArt),1,"Art"), EXPAND("Books",ROWS(tblBooks),1,"Books"), EXPAND("Items",ROWS(tblItems),1,"Items"), EXPAND("Misc",ROWS(tblMisc),1,"Misc")),
 txt, BYROW(CHOOSECOLS(d,1,2,3,4,6,12,15,16), LAMBDA(r, TEXTJOIN(" ",TRUE,r))),
 val, IFERROR(--CHOOSECOLS(d,9),0),
 keep, ((CHOOSECOLS(d,1)<>"")+(CHOOSECOLS(d,2)<>"")>0)
   * IF($B$2="",1,ISNUMBER(SEARCH($B$2,txt)))
   * IF($B$3="",1,cat=$B$3)
   * IF($B$4="",1,CHOOSECOLS(d,14)=$B$4)
   * IF($B$5="",1,ISNUMBER(SEARCH($B$5,CHOOSECOLS(d,12))))
   * IF($B$6="",1,val>=$B$6)
   * IF($B$7="",1,val<=$B$7),
 res, FILTER(HSTACK(cat,d), keep, "No matches"),
 IFERROR(SORT(res,3), res))
'@
    $search.Columns.Item('A').ColumnWidth = 20
    $search.Columns.Item('B:Q').ColumnWidth = 16
    $search.Columns.Item('C').ColumnWidth = 30
    $search.Range('H11:I1000').NumberFormat = '$#,##0.00'
    $search.Range('G11:G1000').NumberFormat = 'yyyy-mm-dd'
    $search.Range('L11:L1000').NumberFormat = 'yyyy-mm-dd'

    # How-to tab.
    $lines = @(
        'COLLECTION INVENTORY - HOW TO USE THIS WORKBOOK',
        '',
        'ADDING ITEMS',
        '  * Type into the table on the Art, Books, Items or Misc tab: start on the first empty row under the table and it grows automatically.',
        '  * Or use the web form, which writes into the same tables.',
        '  * Every column is optional. Item No: copy the next number from the Lists tab (Next Item No), or leave it blank and use "Tidy workbook" in the web form.',
        '',
        'ISBN LOOKUP (books)',
        '  * Type the ISBN on the Books tab, then Data > Refresh All (Excel desktop). Details from Open Library appear on the ISBN Lookup tab.',
        '  * First refresh only: if Excel asks about privacy levels, set both sources to Public, or File > Options > Trust Center > Privacy levels: Ignore for this workbook.',
        '  * The web form "Tidy workbook" action copies looked-up details into blank cells on the Books tab.',
        '',
        'SEARCH',
        '  * Search tab: type a word in B2 (and optional filters in B3:B7). Matching items from every category list below row 10.',
        '',
        'MOVING AN ITEM TO ANOTHER CATEGORY (e.g. out of Misc)',
        '  * Cut the row, paste it as a new row at the bottom of the other table, clear Item No and assign the new number. Note the old number in Notes.',
        '',
        'ADDING A CATEGORY',
        '  1. Right-click the _Template tab > Move or Copy > Create a copy. Rename the tab (e.g. Coins).',
        '  2. Table Design > Table Name: rename the table (e.g. tblCoins). Add the category''s own columns after Last Updated.',
        '  3. Lists tab: add a row to tblCategories (Name, Tab, Table, Prefix).',
        '  4. Search tab, cell A11: add tblCoins[[Item No]:[Notes]] to the first VSTACK and EXPAND("Coins",ROWS(tblCoins),1,"Coins") to the second.',
        '  5. Dropdowns for new columns: add a list table on the Lists tab and a row in tblDropdowns.',
        '',
        'DO NOT rename the shared columns (Item No through Last Updated) or the tables; the web form and Search rely on them.'
    )
    for ($i = 0; $i -lt $lines.Count; $i++) { $howTo.Cells.Item($i + 1, 1).Value2 = $lines[$i] }
    $howTo.Range('A1').Font.Bold = $true
    $howTo.Range('A1').Font.Size = 14
    foreach ($r in 3, 8, 13, 16, 19) { $howTo.Cells.Item($r, 1).Font.Bold = $true }
    $howTo.Columns.Item('A').ColumnWidth = 140

    $howTo.Activate()
    # The template is published; keep the author's name out of the document properties.
    $wb.RemovePersonalInformation = $true
    if (Test-Path $OutFile) { Remove-Item $OutFile -Force }
    $wb.SaveAs($OutFile, 51)   # xlOpenXMLWorkbook (.xlsx)
    $wb.Close($false)
    Write-Output "Saved $OutFile"
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
