#Requires -Version 5
# Opens the workbook, runs one non-looping frame, checks the render buffer.
$ErrorActionPreference = 'Stop'
$book = Join-Path (Split-Path -Parent $PSScriptRoot) 'gauntlex.xlsm'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$wb = $null
try {
    $wb = $excel.Workbooks.Open($book)
    $excel.Run('GameInit')
    $excel.Run('RenderInit')
    $excel.Run('RenderFrame', [double]0)

    $ws = $wb.Worksheets.Item('Screen')
    $topLeft = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1,40)).Value2
    $row1 = -join (1..40 | ForEach-Object { [string]$topLeft.GetValue(1, $_) })

    # find the '@'
    $found = $null
    for ($r = 1; $r -le 25 -and -not $found; $r++) {
        for ($c = 1; $c -le 40; $c++) {
            if ([string]$ws.Cells.Item($r,$c).Value2 -eq '@') { $found = "$r,$c"; break }
        }
    }
    $w = [math]::Round($ws.Columns(1).Width, 2)
    $h = [math]::Round($ws.Rows(1).Height, 2)

    Write-Host "row 1      : $row1"
    Write-Host "player '@' : $(if($found){$found}else{'NOT FOUND'})"
    Write-Host "cell w x h : $w x $h  (want ~18 x 18)"
    Write-Host "hud        : $([string]$ws.Cells.Item(27,1).Value2)"
    if (-not $found) { throw "player token not rendered" }
    Write-Host "OK"
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
