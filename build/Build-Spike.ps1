#Requires -Version 5
<#
  Build spike.xlsm from spike\*.bas (throwaway rendering spike).
  Prereq: build\Enable-VBOM.ps1 once; reference\genesis-tiles\ populated.
#>
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [string]$Out  = (Join-Path (Split-Path -Parent $PSScriptRoot) 'spike.xlsm')
)
$ErrorActionPreference = 'Stop'
$xlMacroEnabled = 52
$ctStd = 1

if (-not (Test-Path (Join-Path $Root 'reference\genesis-tiles\r09_c03.png'))) {
    Write-Host "slicing tiles..."
    & (Join-Path $Root 'reference\slice-tiles.ps1')
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
$tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('gxs_' + [Guid]::NewGuid().ToString('N'))
try {
    if ($null -eq $excel.VBE) { throw "run build\Enable-VBOM.ps1 first" }
    if (Test-Path $Out) {
        $wb = $excel.Workbooks.Open($Out)
        if ($wb.ReadOnly) { throw "spike.xlsm open elsewhere - close it" }
    } else {
        $wb = $excel.Workbooks.Add()
    }
    New-Item -ItemType Directory -Path $tmpDir | Out-Null
    Get-ChildItem (Join-Path $Root 'spike') -Filter '*.bas' | ForEach-Object {
        $name = $_.BaseName
        $text = [IO.File]::ReadAllText($_.FullName) -replace "`r`n", "`n" -replace "`n", "`r`n"
        $comp = $null
        foreach ($c in $wb.VBProject.VBComponents) { if ($c.Name -eq $name) { $comp = $c; break } }
        if ($comp) {
            $body = ($text -split "`r`n" | Where-Object { $_ -notmatch '^Attribute\s+VB_' }) -join "`r`n"
            $cm = $comp.CodeModule
            if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
            $cm.AddFromString($body)
        } else {
            $tmp = Join-Path $tmpDir $_.Name
            [IO.File]::WriteAllText($tmp, $text, (New-Object Text.UTF8Encoding($false)))
            $wb.VBProject.VBComponents.Import($tmp) | Out-Null
        }
        Write-Host "  module $name"
    }
    # Spike sheet + test buttons
    $ws = $null
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq 'Spike') { $ws = $s } }
    if ($null -eq $ws) { $ws = $wb.Worksheets.Add(); $ws.Name = 'Spike' }
    foreach ($sh in @($ws.Shapes)) { if ($sh.Name -like 'btn*') { $sh.Delete() } }
    $tests = @('Test1_Background','Test2_Shapes','Test3_ShapesAnim','Test4_Combined','Test5_ShapeMaze')
    for ($i = 0; $i -lt $tests.Count; $i++) {
        $b = $ws.Shapes.AddShape(1, 6 + $i * 150, 20 * 22 + 30, 140, 30)  # below the VR=20 rows
        $b.Name = 'btn' + $i
        $b.TextFrame2.TextRange.Text = $tests[$i]
        $b.OnAction = $tests[$i]
    }

    if (Test-Path $Out) { $wb.Save() } else { $wb.SaveAs($Out, $xlMacroEnabled) }
    if (-not $wb.Saved) { throw "save did not persist" }
    Write-Host "Built $Out"
}
finally {
    if (Test-Path $tmpDir) { Remove-Item -Recurse -Force $tmpDir }
    if ($wb) { try { $wb.Close($false) } catch {} }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
