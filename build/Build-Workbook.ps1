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

$MAP_ROWS  = 32          # full level grid stamped onto each level sheet
$MAP_COLS  = 32

& (Join-Path $PSScriptRoot 'Check-Level.ps1')   # fail the build on an unbeatable level

# the picture-Shape renderer reads reference\genesis-tiles\*.png at runtime
if (-not (Test-Path (Join-Path $Root 'reference\genesis-tiles\r13_c11.png'))) {
    & (Join-Path $Root 'reference\slice-tiles.ps1')
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
try {
    try { $probe = $excel.VBE } catch {}
    if ($null -eq $excel.VBE) { throw "Cannot reach the VBA project model. Run build\Enable-VBOM.ps1 first." }

    if (Test-Path $Out) {
        $wb = $excel.Workbooks.Open($Out)
        if ($wb.ReadOnly) {
            throw "gauntlex.xlsm opened read-only - it is open in another Excel window. Close it and re-run."
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
    # Anchored at A1 (top-left of the viewport). StartGauntlex hides it while
    # the loop runs and shows it again on exit, so it doubles as "restart".
    $btn = $screen.Shapes.AddShape($msoShapeRectangle, 2, 2, 150, 28)
    $btn.Name = 'btnPlay'
    $btn.Placement = 3          # xlFreeFloating - don't move/size with cells
    $btn.TextFrame2.TextRange.Text = "PLAY"
    $btn.OnAction = "StartGauntlex"

    # ---- Levels ----
    Get-ChildItem (Join-Path $Root 'levels') -Filter '*.txt' | ForEach-Object {
        $name  = $_.BaseName
        $sheet = Ensure-Sheet $name
        $sheet.Cells.Clear() | Out-Null
        $lines = Get-Content -LiteralPath $_.FullName
        for ($r = 0; $r -lt $MAP_ROWS; $r++) {
            $line = if ($r -lt $lines.Count) { [string]$lines[$r] } else { '' }
            for ($c = 0; $c -lt $MAP_COLS; $c++) {
                $ch = if ($c -lt $line.Length) { [string]$line[$c] } else { '.' }
                if ($ch -eq ' ') { $ch = '.' }
                $sheet.Cells.Item($r + 1, $c + 1).Value2 = $ch
            }
        }
        $sheet.Visible = $xlVeryHidden
        Write-Host "  level $name  <- $($_.Name)"
    }

    # ---- VBA modules ----
    # Overwrite code in place when the module already exists: Remove + re-Import
    # of a same-named module in one session does not commit (the import is
    # dropped, the stale module kept). VBE also needs CRLF, not LF.
    $srcFiles = Get-ChildItem (Join-Path $Root 'src') -Include '*.bas','*.cls' -File -Recurse
    $srcNames = @()
    $tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('gx_' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmpDir | Out-Null
    try {
        foreach ($f in $srcFiles) {
            $name = [IO.Path]::GetFileNameWithoutExtension($f.FullName)
            $srcNames += $name
            $text = [IO.File]::ReadAllText($f.FullName) -replace "`r`n", "`n" -replace "`n", "`r`n"

            $comp = $null
            foreach ($c in $wb.VBProject.VBComponents) { if ($c.Name -eq $name) { $comp = $c; break } }

            if ($comp -and $f.Extension -eq '.bas') {
                $body = ($text -split "`r`n" | Where-Object { $_ -notmatch '^Attribute\s+VB_' }) -join "`r`n"
                $cm = $comp.CodeModule
                if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
                $cm.AddFromString($body)
                Write-Host "  module $name  (replaced)"
            } else {
                if ($comp) { $wb.VBProject.VBComponents.Remove($comp) }
                $tmp = Join-Path $tmpDir $f.Name
                [IO.File]::WriteAllText($tmp, $text, (New-Object Text.UTF8Encoding($false)))
                $wb.VBProject.VBComponents.Import($tmp) | Out-Null
                Write-Host "  module $name  (imported)"
            }
        }
    } finally {
        Remove-Item -Recurse -Force $tmpDir
    }

    # drop std/class modules that no longer have a source file
    foreach ($c in @($wb.VBProject.VBComponents)) {
        if (($c.Type -eq $ctStdModule -or $c.Type -eq $ctClassModule) -and ($srcNames -notcontains $c.Name)) {
            Write-Host "  module $($c.Name)  (removed - no source)"
            $wb.VBProject.VBComponents.Remove($c)
        }
    }

    # ---- drop any stray default sheets ----
    foreach ($s in @($wb.Worksheets)) {
        if ($s.Name -ne 'Screen' -and $s.Name -notmatch '^L\d+$' -and $wb.Worksheets.Count -gt 1) {
            $s.Delete()
        }
    }

    # sanity: catch a stale / truncated module import before we ship the workbook
    $gm = $wb.VBProject.VBComponents('modGame').CodeModule
    if ($gm.Lines(1, $gm.CountOfLines) -notmatch 'Sub DebugStep') {
        throw "modGame looks stale (no DebugStep) - module replace failed"
    }

    $screen.Activate()
    if (Test-Path $Out) { $wb.Save() } else { $wb.SaveAs($Out, $xlMacroEnabled) }
    if (-not $wb.Saved) { throw "save did not persist - is gauntlex.xlsm open elsewhere?" }
    Write-Host "Built $Out"
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
