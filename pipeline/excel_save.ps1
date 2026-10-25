# Opens the Excel dashboard in the installed Excel and saves it once: Excel then stores the chart caches and
# calculated values that Excel for the web shows. Called by pipeline/build_excel_dashboard.py.
param(
    [Parameter(Mandatory = $true)][string]$Path
)
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $Path).Path)
    [void]$wb.Worksheets.Item(1).Activate()
    [void]$wb.Worksheets.Item(1).Range('A1').Select()
    $xl.CalculateFull()
    $wb.Save()
    Write-Output "Saved with Excel: $Path"
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
