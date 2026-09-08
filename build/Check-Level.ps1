#Requires -Version 5
<#
  Static reachability check for levels\*.txt (block grid).
  Flood-fills from the spawn and verifies keys/food/grunts/exit are actually
  reachable - catches sealed rooms the Excel smoke test can't (it teleports).

  Usage: powershell -ExecutionPolicy Bypass -File build\Check-Level.ps1
#>
param([string]$Root = (Split-Path -Parent $PSScriptRoot))
$ErrorActionPreference = 'Stop'
$W = 32; $H = 32
$fail = 0

function Flood($g, $sr, $sc, $throughDoors) {
    $seen = New-Object 'bool[,]' $H, $W
    $q = [System.Collections.Generic.Queue[int[]]]::new()
    $q.Enqueue(@($sr, $sc)); $seen[$sr, $sc] = $true
    while ($q.Count) {
        $p = $q.Dequeue()
        foreach ($d in @(@(-1,0),@(1,0),@(0,-1),@(0,1))) {
            $r = $p[0] + $d[0]; $c = $p[1] + $d[1]
            if ($r -lt 0 -or $r -ge $H -or $c -lt 0 -or $c -ge $W) { continue }
            if ($seen[$r, $c]) { continue }
            $ch = $g[$r][$c]
            if ($ch -eq '#' -or 'GOEZL'.Contains([string]$ch)) { continue }   # generators are solid
            if ($ch -eq 'D' -and -not $throughDoors) { continue }
            $seen[$r, $c] = $true
            $q.Enqueue(@($r, $c))
        }
    }
    $seen
}

Get-ChildItem (Join-Path $Root 'levels') -Filter '*.txt' | ForEach-Object {
    $name = $_.BaseName
    $lines = Get-Content -LiteralPath $_.FullName
    $g = @(); for ($r = 0; $r -lt $H; $r++) {
        $ln = if ($r -lt $lines.Count) { [string]$lines[$r] } else { '' }
        $ln = $ln.PadRight($W).Substring(0, $W)
        $g += , ($ln.ToCharArray())
    }

    $spawn = $null; $cells = @{}
    for ($r = 0; $r -lt $H; $r++) { for ($c = 0; $c -lt $W; $c++) {
        $ch = [string]$g[$r][$c]
        if ($ch -eq 'S') { $spawn = @($r, $c) }
        if ('KDX+PGOEZL'.Contains($ch)) { $cells[$ch] += , @($r, $c) }
    } }

    Write-Host "== $name =="
    if (-not $spawn) { Write-Host "  FAIL no spawn (S)"; $script:fail++; return }

    $reachNoDoor = Flood $g $spawn[0] $spawn[1] $false
    $reachDoor   = Flood $g $spawn[0] $spawn[1] $true

    function Report($ch, $label, $set) {
        $pts = $cells[$ch]
        if (-not $pts) { Write-Host "  --   no $label"; return }
        foreach ($p in $pts) {
            $ok = $set[$p[0], $p[1]]
            if ($ok) { Write-Host "  ok   $label ($($p[0]),$($p[1]))" }
            else     { Write-Host "  FAIL $label ($($p[0]),$($p[1])) unreachable"; $script:fail++ }
        }
    }
    # keys / food / potions must be reachable WITHOUT opening doors (no key yet)
    Report 'K' 'key'    $reachNoDoor
    Report '+' 'food'   $reachNoDoor
    Report 'P' 'potion' $reachNoDoor
    # exit must be reachable once doors can be opened
    Report 'X' 'exit'  $reachDoor

    # doors/generators are solid: each must border reachable floor.
    # a door must be reachable BEFORE any door opens (so you can use the key on it).
    function Approachable($ch, $label, $set) {
        foreach ($p in $cells[$ch]) {
            $adj = $false
            foreach ($d in @(@(-1,0),@(1,0),@(0,-1),@(0,1))) {
                $r = $p[0]+$d[0]; $c = $p[1]+$d[1]
                if ($r -ge 0 -and $r -lt $H -and $c -ge 0 -and $c -lt $W -and $set[$r,$c]) { $adj = $true }
            }
            if ($adj) { Write-Host "  ok   $label ($($p[0]),$($p[1])) approachable" }
            else      { Write-Host "  FAIL $label ($($p[0]),$($p[1])) not approachable"; $script:fail++ }
        }
    }
    Approachable 'D' 'door'       $reachNoDoor
    Approachable 'G' 'grunt-gen'  $reachDoor
    Approachable 'O' 'ghost-gen'  $reachDoor
    Approachable 'E' 'demon-gen'  $reachDoor
    Approachable 'Z' 'sorc-gen'   $reachDoor
    Approachable 'L' 'lobber-gen' $reachDoor
}

if ($fail) { throw "$fail level check(s) failed" }
Write-Host "levels OK"
