#Requires -Version 5
# Opens the workbook, drives it headless, checks render + M3a game logic.
$ErrorActionPreference = 'Stop'
$book = Join-Path (Split-Path -Parent $PSScriptRoot) 'gauntlex.xlsm'

$VC = 36; $VR = 20
$START_HEALTH = 2000; $START_LIVES = 3
$K_GRUNT = 1; $K_GHOST = 2; $K_DEMON = 3
$GHOST_DMG = 55; $DEMON_SHOT_DMG = 70; $SCORE_GRUNT = 10; $SCORE_GEN = 100

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false
$wb = $null; $fail = 0
function Check($label, $cond, $got) {
    if ($cond) { Write-Host "  ok   $label" }
    else       { Write-Host "  FAIL $label  (got: $got)"; $script:fail++ }
}
try {
    $wb = $excel.Workbooks.Open($book)
    $ws = $wb.Worksheets.Item('Screen')

    function View-Row($r) {
        $v = $ws.Range($ws.Cells.Item($r,1), $ws.Cells.Item($r,$VC)).Value2
        -join (1..$VC | ForEach-Object { [string]$v.GetValue(1, $_) })
    }
    function Find($glyph) {
        $first = $null; $n = 0
        for ($r = 1; $r -le $VR; $r++) { for ($c = 1; $c -le $VC; $c++) {
            if ([string]$ws.Cells.Item($r,$c).Value2 -eq $glyph) { if (-not $first) { $first = "$r,$c" }; $n++ }
        } }
        [pscustomobject]@{ first = $first; n = $n }
    }
    # DebugState -> vc;vr;camR;camC;plHR;plHC;state;cell;lives;health;keys;score;entN;e1r;e1c;genAlive;prjN
    function State {
        $p = ([string]$excel.Run('DebugState')).Split(';')
        [pscustomobject]@{
            vc=[int]$p[0]; vr=[int]$p[1]; camR=[int]$p[2]; camC=[int]$p[3]
            plHR=[int]$p[4]; plHC=[int]$p[5]; st=$p[6]; cell=[double]$p[7]
            lives=[int]$p[8]; health=[int]$p[9]; keys=[int]$p[10]; score=[int]$p[11]
            entN=[int]$p[12]; e1r=[int]$p[13]; e1c=[int]$p[14]; genAlive=[int]$p[15]; prjN=[int]$p[16]
        }
    }
    function Dist($s) { [math]::Abs($s.e1r - $s.plHR) + [math]::Abs($s.e1c - $s.plHC) }
    function Far { $excel.Run('GameInit'); $excel.Run('DebugWarp', [int]57, [int]5) }   # far corner, no generators near

    # ---- render / camera (M1) ----
    $excel.Run('GameInit'); $excel.Run('FitViewport'); $excel.Run('RenderInit')
    $excel.Run('CenterCamera'); $excel.Run('RenderFrame', [double]0)
    $s = State
    Check "viewport 36x20"           ($s.vc -eq $VC -and $s.vr -eq $VR) "$($s.vc)x$($s.vr)"
    $uh=[double]$excel.ActiveWindow.UsableHeight; $uw=[double]$excel.ActiveWindow.UsableWidth
    $exp=[math]::Max(8,[math]::Min(60,[math]::Min($uh/$VR,($uw-220)/$VC)))
    Check "cell fits window"         ([math]::Abs($s.cell-$exp) -le 1.0) "$($s.cell) vs ~$([math]::Round($exp,1))"
    Check "camera clamped 1,1"       ($s.camR -eq 1 -and $s.camC -eq 1) "$($s.camR),$($s.camC)"
    Check "top-left 2x2 wall"        ((View-Row 1).Substring(0,2) -eq '##') (View-Row 1)
    Check "player 2x2 at view 3,3"   ((Find '@').n -eq 4 -and (Find '@').first -eq '3,3') (Find '@').first
    Check "HUD shows LIVES"          ([string]$ws.Cells.Item(6, $VC+2).Value2 -like 'LIVES*') ([string]$ws.Cells.Item(6, $VC+2).Value2)

    $excel.Run('DebugStep', [int]14, [int]30)
    $excel.Run('RenderFrame', [double]0); $s = State
    Check "scroll: player 17,33 cam 7,15" ($s.plHR -eq 17 -and $s.plHC -eq 33 -and $s.camR -eq 7 -and $s.camC -eq 15) "$($s.plHR),$($s.plHC)/$($s.camR),$($s.camC)"
    Check "player re-rendered view 11,19"  ((Find '@').first -eq '11,19') (Find '@').first

    # ---- M3a: level starts with 4 generators, no live entities ----
    $excel.Run('GameInit'); $s = State
    Check "4 generators, 0 entities"  ($s.genAlive -eq 4 -and $s.entN -eq 0) "gen $($s.genAlive) ent $($s.entN)"
    Check "start health/lives"        ($s.health -eq $START_HEALTH -and $s.lives -eq $START_LIVES) "$($s.health)/$($s.lives)"

    # ---- M3a: generators spawn while on screen ----
    $excel.Run('GameInit')
    $excel.Run('DebugUpdate', [int]3600)
    $s = State
    Check "generator spawned entities" ($s.entN -ge 1) $s.entN
    $excel.Run('RenderFrame', [double]0)
    Check "an entity glyph renders"     ((Find 'g').n + (Find 'o').n + (Find 'd').n -ge 1) "g/o/d present"

    # ---- M3a: health drains (far from generators, none spawn) ----
    Far
    $excel.Run('DebugUpdate', [int]2200)
    Check "health drained ~2"        ((State).health -le $START_HEALTH-2 -and (State).health -ge $START_HEALTH-4) (State).health

    # ---- M3a: a spawned grunt pursues the player ----
    Far
    $excel.Run('DebugSpawn', [int]$K_GRUNT, [int]57, [int]25)
    $d0 = Dist (State); $excel.Run('DebugUpdate', [int]800); $d1 = Dist (State)
    Check "grunt closed distance"     ($d1 -lt $d0) "before $d0 after $d1"

    # ---- M3a: walking into a grunt melee-kills it ----
    Far
    $excel.Run('DebugSpawn', [int]$K_GRUNT, [int]57, [int]7)
    $b = State
    $excel.Run('DebugStep', [int]0, [int]1)
    $a = State
    Check "melee killed grunt +10"    ($a.entN -eq 0 -and $a.score -eq $b.score + $SCORE_GRUNT) "ent $($a.entN) score $($b.score)->$($a.score)"

    # ---- M3a: player shot kills a grunt down-range ----
    Far
    $excel.Run('DebugFace', [int]0, [int]1)
    $excel.Run('DebugSpawn', [int]$K_GRUNT, [int]57, [int]22)
    $excel.Run('DebugFire')
    $excel.Run('DebugUpdate', [int]900)
    $s = State
    Check "shot killed grunt +10"     ($s.entN -eq 0 -and $s.score -ge $SCORE_GRUNT) "ent $($s.entN) score $($s.score)"

    # ---- M3a: ghost is a kamikaze ----
    Far
    $hp0 = (State).health
    $excel.Run('DebugSpawn', [int]$K_GHOST, [int]58, [int]6)
    $excel.Run('DebugUpdate', [int]120)
    $s = State
    Check "ghost kamikaze hit + died" ($s.entN -eq 0 -and $s.health -eq $hp0 - $GHOST_DMG) "ent $($s.entN) hp $hp0->$($s.health)"

    # ---- M3a: demon fires (shot lands on the in-line player) ----
    Far
    $hp0 = (State).health
    $excel.Run('DebugSpawn', [int]$K_DEMON, [int]57, [int]40)
    $excel.Run('DebugUpdate', [int]1500)
    Check "demon shot hit player"     ((State).health -le $hp0 - $DEMON_SHOT_DMG) "$hp0 -> $((State).health)"

    # ---- M3a: a generator is destroyed by repeated melee ----
    $excel.Run('GameInit')
    $excel.Run('DebugWarp', [int]17, [int]13)          # below grunt-gen at block (8,7)
    $b = State
    1..3 | ForEach-Object { $excel.Run('DebugStep', [int]-1, [int]0) }
    $a = State
    Check "generator destroyed +100"  ($a.genAlive -eq 3 -and $a.score -eq $b.score + $SCORE_GEN) "gen $($b.genAlive)->$($a.genAlive) score +$($a.score - $b.score)"

    # ---- M2 carry-over: key / door / exit / death ----
    $excel.Run('GameInit'); $excel.Run('DebugWarp', [int]9, [int]43); $excel.Run('DebugStep', [int]0, [int]2)
    Check "key picked up"            ((State).keys -eq 1) (State).keys

    $excel.Run('GameInit'); $excel.Run('DebugWarp', [int]51, [int]57); $excel.Run('DebugStep', [int]2, [int]0)
    Check "door blocks without key"  ((State).plHR -eq 51) (State).plHR
    $excel.Run('DebugWarp', [int]9, [int]43); $excel.Run('DebugStep', [int]0, [int]2)
    $excel.Run('DebugWarp', [int]51, [int]57); $excel.Run('DebugStep', [int]2, [int]0)
    $s = State
    Check "door opens with key"      ($s.plHR -eq 53 -and $s.keys -eq 0) "plHR $($s.plHR) keys $($s.keys)"

    $excel.Run('GameInit'); $excel.Run('DebugWarp', [int]57, [int]57); $excel.Run('DebugUpdate', [int]50)
    Check "exit -> WON +100"          ((State).st -eq 'WON' -and (State).score -eq 100) "$((State).st) $((State).score)"

    $excel.Run('GameInit'); $excel.Run('DebugSetHealth', [int]2); $excel.Run('DebugUpdate', [int]3400)
    $s = State
    Check "death: life lost + respawn" ($s.lives -eq 2 -and $s.health -eq $START_HEALTH -and $s.st -eq 'PLAY' -and $s.plHR -eq 3) "lives $($s.lives) hp $($s.health) $($s.st)"
    $excel.Run('DebugSetHealth', [int]1); $excel.Run('DebugUpdate', [int]1200)
    $excel.Run('DebugSetHealth', [int]1); $excel.Run('DebugUpdate', [int]1200)
    Check "game over at 0 lives"      ((State).lives -eq 0 -and (State).st -eq 'OVER') "$((State).lives) $((State).st)"

    if ($fail) { throw "$fail check(s) failed" }
    Write-Host "OK"
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
