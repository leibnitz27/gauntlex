#Requires -Version 5
<#
    Enables "Trust access to the VBA project object model" for the current user.
    Without this, Build-Workbook.ps1 cannot import modules (COM throws
    "Programmatic access to Visual Basic Project is not trusted").

    This writes one HKCU DWORD:
      HKCU\Software\Microsoft\Office\<ver>\Excel\Security\AccessVBOM = 1
    Reverse it by setting the value back to 0.
#>
$ErrorActionPreference = 'Stop'

$excel = New-Object -ComObject Excel.Application
$ver = $excel.Version            # e.g. "16.0"
$excel.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)

$key = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name AccessVBOM -Value 1 -Type DWord
Write-Host "AccessVBOM = 1  (Office $ver).  You can now run build\Build-Workbook.ps1."
