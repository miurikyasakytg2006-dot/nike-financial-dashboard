# Adds the year and growth-basis slicers to the Excel dashboard (XlsxWriter cannot create slicers),
# then saves the workbook with Excel, which also stores calculated formula results.
# Called by pipeline/build_excel_dashboard.py. Anchors are "sheetIndex:cell:widthPt:heightPt";
# several year anchors are comma-separated and all share one slicer cache, so they stay in sync.
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$YearAnchors,
    [Parameter(Mandatory = $true)][string]$BasisAnchor
)
$ErrorActionPreference = 'Stop'
$missing = [Type]::Missing
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $Path).Path)
    $data = $wb.Worksheets.Item($wb.Worksheets.Count)

    function Add-Slicers($table, $anchors, $columns, $sort) {
        $lo = $data.ListObjects.Item($table)
        $cache = $wb.SlicerCaches.Add2($lo, $lo.ListColumns.Item(1).Name)
        $cache.SortItems = $sort
        $i = 0
        foreach ($a in $anchors.Split(',')) {
            $p = $a.Split(':')
            $ws = $wb.Worksheets.Item([int]$p[0])
            $cell = $ws.Range($p[1])
            $i++
            $s = $cache.Slicers.Add($ws, $missing, "sl_$table$i", $missing, $cell.Top, $cell.Left, [double]$p[2], [double]$p[3])
            $s.NumberOfColumns = $columns
            $s.Style = 'SlicerStyleLight1'
            $s.DisplayHeader = $true
        }
    }
    # Table slicers sort only by value: 2 = ascending (FY2020..FY2026), 3 = descending ("Theo bao cao" before "Loai tru ty gia")
    Add-Slicers 'tblNam' $YearAnchors 7 2
    Add-Slicers 'tblCoSo' $BasisAnchor 2 3

    [void]$wb.Worksheets.Item(1).Activate()
    [void]$wb.Worksheets.Item(1).Range('A1').Select()
    $xl.CalculateFull()
    $wb.Save()
    Write-Output "Slicers added: $Path"
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
