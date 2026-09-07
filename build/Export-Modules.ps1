#Requires -Version 5
<#
    Export VBA modules from gauntlex.xlsm back to src\ as text.
    Run this after editing code in the VBA IDE so git sees the change.

    Prereq: run build\Enable-VBOM.ps1 once.
    Usage:  powershell -ExecutionPolicy Bypass -File build\Export-Modules.ps1
#>
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [string]$Book = (Join-Path (Split-Path -Parent $PSScriptRoot) 'gauntlex.xlsm')
)
$ErrorActionPreference = 'Stop'

$ctStdModule   = 1
$ctClassModule = 2
$src = Join-Path $Root 'src'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
try {
    if ($null -eq $excel.VBE) { throw "Cannot reach the VBA project model. Run build\Enable-VBOM.ps1 first." }
    $wb = $excel.Workbooks.Open($Book)
    foreach ($comp in $wb.VBProject.VBComponents) {
        switch ($comp.Type) {
            $ctStdModule   { $ext = 'bas' }
            $ctClassModule { $ext = 'cls' }
            default        { $ext = $null }
        }
        if ($ext) {
            $path = Join-Path $src ("{0}.{1}" -f $comp.Name, $ext)
            $comp.Export($path)
            Write-Host "  $($comp.Name).$ext"
        }
    }
    Write-Host "Exported to $src"
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
