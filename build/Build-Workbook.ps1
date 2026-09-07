#Requires -Version 5
<#
    Build gauntlex.xlsm from source:
      - (re)imports every src\*.bas / *.cls module
      - stamps each levels\*.txt onto a matching very-hidden sheet
      - ensures the "Screen" sheet + a PLAY button

    Prereq: run build\Enable-VBOM.ps1 once (see that file).
    Usage:  powershell -ExecutionPolicy Bypass -File build\Build-Workbook.ps1
#>
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [string]$Out  = (Join-Path (Split-Path -Parent $PSScriptRoot) 'gauntlex.xlsm')
)
$ErrorActionPreference = 'Stop'

$xlMacroEnabled   = 52     # xlOpenXMLWorkbookMacroEnabled
$ctStdModule      = 1
$ctClassModule    = 2
$xlVeryHidden     = 2
$msoShapeRectangle = 1

$VIEW_ROWS = 25
$VIEW_COLS = 40
$CELL_PTS  = 18

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
try {
    try { $probe = $excel.VBE } catch {}
    if ($null -eq $excel.VBE) { throw "Cannot reach the VBA project model. Run build\Enable-VBOM.ps1 first." }

    if (Test-Path $Out) {
        $wb = $excel.Workbooks.Open($Out)
        foreach ($comp in @($wb.VBProject.VBComponents)) {
            if ($comp.Type -eq $ctStdModule -or $comp.Type -eq $ctClassModule) {
                $wb.VBProject.VBComponents.Remove($comp)
            }
        }
    } else {
        $wb = $excel.Workbooks.Add()
    }

    function Ensure-Sheet($name) {
        foreach ($s in $wb.Worksheets) { if ($s.Name -eq $name) { return $s } }
        $s = $wb.Worksheets.Add()
        $s.Name = $name
        return $s
    }

    # ---- Screen sheet ----
    $screen = Ensure-Sheet 'Screen'
    $screen.Visible = -1
    $screen.Cells.Clear() | Out-Null
    foreach ($sh in @($screen.Shapes)) { $sh.Delete() }
    $btn = $screen.Shapes.AddShape($msoShapeRectangle, 6, ($VIEW_ROWS + 4) * $CELL_PTS, 160, 30)
    $btn.TextFrame2.TextRange.Text = "PLAY  >  StartGauntlex"
    $btn.OnAction = "StartGauntlex"

    # ---- Levels ----
    Get-ChildItem (Join-Path $Root 'levels') -Filter '*.txt' | ForEach-Object {
        $name  = $_.BaseName
        $sheet = Ensure-Sheet $name
        $sheet.Cells.Clear() | Out-Null
        $lines = Get-Content -LiteralPath $_.FullName
        for ($r = 0; $r -lt $VIEW_ROWS; $r++) {
            $line = if ($r -lt $lines.Count) { [string]$lines[$r] } else { '' }
            for ($c = 0; $c -lt $VIEW_COLS; $c++) {
                $ch = if ($c -lt $line.Length) { [string]$line[$c] } else { '.' }
                if ($ch -eq ' ') { $ch = '.' }
                $sheet.Cells.Item($r + 1, $c + 1).Value2 = $ch
            }
        }
        $sheet.Visible = $xlVeryHidden
        Write-Host "  level $name  <- $($_.Name)"
    }

    # ---- VBA modules ----
    Get-ChildItem (Join-Path $Root 'src') -Include '*.bas','*.cls' -File -Recurse | ForEach-Object {
        $wb.VBProject.VBComponents.Import($_.FullName) | Out-Null
        Write-Host "  module $($_.BaseName)"
    }

    # ---- drop any stray default sheets ----
    foreach ($s in @($wb.Worksheets)) {
        if ($s.Name -ne 'Screen' -and $s.Name -notmatch '^L\d+$' -and $wb.Worksheets.Count -gt 1) {
            $s.Delete()
        }
    }

    $screen.Activate()
    if (Test-Path $Out) { $wb.Save() } else { $wb.SaveAs($Out, $xlMacroEnabled) }
    Write-Host "Built $Out"
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
