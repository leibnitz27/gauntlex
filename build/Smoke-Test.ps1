#Requires -Version 5
# Opens the workbook, drives it headless, checks render + M3 game logic.
$ErrorActionPreference = 'Stop'
$book = Join-Path (Split-Path -Parent $PSScriptRoot) 'gauntlex.xlsm'

$VC = 36; $VR = 20
$START_HEALTH = 2000; $START_LIVES = 3
$K_GRUNT=1; $K_GHOST=2; $K_DEMON=3; $K_SORC=4; $K_LOBBER=5; $K_THIEF=6; $K_DEATH=7
$GHOST_DMG=55; $DEMON_SHOT_DMG=70; $LOBBER_DMG=50; $DEATH_DRAIN=7
$SCORE_GRUNT=10; $SCORE_GEN=100; $SCORE_THIEF=200
$CH_WARRIOR=0; $CH_VALKYRIE=1; $CH_WIZARD=2; $CH_ELF=3
$ARM_WARRIOR=90                       # default char for the damage checks; combat dmg *= pct\100
function Dmg($raw) { [math]::Floor($raw * $ARM_WARRIOR / 100) }

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
    $excel.Run('DebugSetChar', [int]$CH_WARRIOR)     # default for all non-M4a checks

    function View-Row($r) {
        $v = $ws.Range($ws.Cells.Item($r,1), $ws.Cells.Item($r,$VC)).Value2
        -join (1..$VC | ForEach-Object { [string]$v.GetValue(1, $_) })
    }
    function Find($glyph) {
        $first=$null; $n=0
        for ($r=1;$r -le $VR;$r++){for($c=1;$c -le $VC;$c++){
            if ([string]$ws.Cells.Item($r,$c).Value2 -eq $glyph){ if(-not $first){$first="$r,$c"}; $n++ }
        }}
        [pscustomobject]@{ first=$first; n=$n }
    }
    # DebugState -> ..;state;cell;lives;health;keys;score;entN;e1r;e1c;genAlive;prjN;potions;e1kind;char
    function State {
        $p = ([string]$excel.Run('DebugState')).Split(';')
        [pscustomobject]@{
            camR=[int]$p[2]; camC=[int]$p[3]; plHR=[int]$p[4]; plHC=[int]$p[5]; st=$p[6]; cell=[double]$p[7]
            lives=[int]$p[8]; health=[int]$p[9]; keys=[int]$p[10]; score=[int]$p[11]
            entN=[int]$p[12]; e1r=[int]$p[13]; e1c=[int]$p[14]; genAlive=[int]$p[15]; prjN=[int]$p[16]
            pot=[int]$p[17]; e1k=[int]$p[18]; char=[int]$p[19]; lvl=[int]$p[20]
        }
    }
    function Dist($s) { [math]::Abs($s.e1r - $s.plHR) + [math]::Abs($s.e1c - $s.plHC) }
    function Far { $excel.Run('GameInit'); $excel.Run('DebugKillGens'); $excel.Run('DebugWarp', [int]57, [int]5) }

    # ---- render / camera (M1) ----
    $excel.Run('GameInit'); $excel.Run('FitViewport'); $excel.Run('RenderInit')
    $excel.Run('CenterCamera'); $excel.Run('RenderFrame', [double]0); $s = State
    Check "viewport / camera clamp"  ($s.camR -eq 1 -and $s.camC -eq 1) "$($s.camR),$($s.camC)"
    $uh=[double]$excel.ActiveWindow.UsableHeight; $uw=[double]$excel.ActiveWindow.UsableWidth
    $exp=[math]::Max(8,[math]::Min(60,[math]::Min($uh/$VR,($uw-220)/$VC)))
    Check "cell fits window"         ([math]::Abs($s.cell-$exp) -le 1.0) "$($s.cell) vs ~$([math]::Round($exp,1))"
    Check "player 2x2 at view 3,3"   ((Find '@').n -eq 4 -and (Find '@').first -eq '3,3') (Find '@').first
    $excel.Run('DebugStep', [int]14, [int]30); $excel.Run('RenderFrame', [double]0); $s = State
    Check "scroll: player 17,33 cam 7,15" ($s.plHR -eq 17 -and $s.plHC -eq 33 -and $s.camR -eq 7 -and $s.camC -eq 15) "$($s.plHR),$($s.plHC)"

    # ---- M3a: generators / fire / grunt / ghost / demon ----
    $excel.Run('GameInit'); $s = State
    Check "6 generators, 0 entities" ($s.genAlive -eq 6 -and $s.entN -eq 0) "gen $($s.genAlive) ent $($s.entN)"
    $excel.Run('GameInit'); $excel.Run('DebugUpdate', [int]3600)
    Check "generators spawned"       ((State).entN -ge 1) (State).entN

    Far; $excel.Run('DebugUpdate', [int]2200)
    Check "health drains ~2"         ((State).health -le $START_HEALTH-2 -and (State).health -ge $START_HEALTH-4) (State).health

    Far; $excel.Run('DebugSpawn', [int]$K_GRUNT, [int]57, [int]25)
    $d0 = Dist (State); $excel.Run('DebugUpdate', [int]800)
    Check "grunt pursues"            ((Dist (State)) -lt $d0) "$d0 -> $(Dist (State))"

    Far; $excel.Run('DebugSpawn', [int]$K_GRUNT, [int]57, [int]7); $b = State
    $excel.Run('DebugStep', [int]0, [int]1); $a = State
    Check "melee kills grunt +10"    ($a.entN -eq 0 -and $a.score -eq $b.score + $SCORE_GRUNT) "ent $($a.entN) +$($a.score-$b.score)"

    Far; $excel.Run('DebugFace', [int]0, [int]1); $excel.Run('DebugSpawn', [int]$K_GRUNT, [int]57, [int]22)
    $excel.Run('DebugFire'); $excel.Run('DebugUpdate', [int]900)
    Check "shot kills grunt"         ((State).entN -eq 0 -and (State).score -ge $SCORE_GRUNT) "ent $((State).entN)"

    Far; $hp0=(State).health; $excel.Run('DebugSpawn', [int]$K_GHOST, [int]58, [int]6); $excel.Run('DebugUpdate', [int]120)
    Check "ghost kamikaze"           ((State).entN -eq 0 -and (State).health -eq $hp0 - (Dmg $GHOST_DMG)) "ent $((State).entN) hp -$($hp0-(State).health)"

    Far; $hp0=(State).health; $excel.Run('DebugSpawn', [int]$K_DEMON, [int]57, [int]40); $excel.Run('DebugUpdate', [int]1500)
    Check "demon shot hits player"   ((State).health -le $hp0 - (Dmg $DEMON_SHOT_DMG)) "-$($hp0-(State).health)"

    $excel.Run('GameInit'); $excel.Run('DebugWarp', [int]17, [int]13); $b = State
    1..3 | ForEach-Object { $excel.Run('DebugStep', [int]-1, [int]0) }
    Check "generator destroyed +100" ((State).genAlive -eq 5 -and (State).score -eq $b.score + $SCORE_GEN) "gen $($b.genAlive)->$((State).genAlive)"

    # ---- M3b: sorcerer flickers ----
    Far; $excel.Run('DebugSpawn', [int]$K_SORC, [int]57, [int]25)
    $seen = @{}
    0..40 | ForEach-Object { $excel.Run('DebugUpdate', [int]70); $seen["$([bool]$excel.Run('DebugSorcVisible'))"] = 1 }
    Check "sorcerer toggles visibility" ($seen.Count -eq 2) "states: $($seen.Keys -join ',')"

    # ---- M3b: lobber rock arcs over a wall ----
    $excel.Run('GameInit'); $excel.Run('DebugWarp', [int]17, [int]13)   # below box A / grunt-gen (8,7)
    $hp0 = (State).health
    $excel.Run('DebugSpawn', [int]$K_LOBBER, [int]9, [int]13)           # inside sealed box A, walls + a generator between
    $excel.Run('DebugUpdate', [int]2000)
    Check "lobber rock crossed wall" ((State).health -le $hp0 - (Dmg $LOBBER_DMG)) "-$($hp0-(State).health)"

    # ---- M3b: potion pickup + screen-clear blast ----
    $excel.Run('GameInit'); $excel.Run('DebugWarp', [int]5, [int]47); $excel.Run('DebugStep', [int]0, [int]2)  # onto P at block (3,25)
    Check "potion picked up"         ((State).pot -eq 1) (State).pot

    $excel.Run('GameInit'); $excel.Run('DebugUpdate', [int]4200)        # let generators fill the view
    $b = State
    $excel.Run('DebugSetPotions', [int]1); $excel.Run('DebugUsePotion')
    $a = State
    Check "potion cleared the view"  ($a.entN -eq 0 -and $a.genAlive -lt $b.genAlive -and $a.pot -eq 0) "ent $($b.entN)->$($a.entN) gen $($b.genAlive)->$($a.genAlive)"

    # ---- M3b: thief steals a key, killing it returns it ----
    $excel.Run('GameInit'); $excel.Run('DebugWarp', [int]57, [int]5); $excel.Run('DebugSetKeys', [int]2)
    $excel.Run('DebugSpawn', [int]$K_THIEF, [int]58, [int]6)
    $excel.Run('DebugUpdate', [int]90)
    $mid = State
    Check "thief stole a key"        ($mid.keys -eq 1) $mid.keys
    $excel.Run('DebugWarp', [int]$mid.e1r, [int]$mid.e1c); $excel.Run('DebugStep', [int]0, [int]0)
    $a = State
    Check "killing thief returns key + score" ($a.keys -eq 2 -and $a.entN -eq 0 -and $a.score -ge $SCORE_THIEF) "keys $($a.keys) ent $($a.entN) score $($a.score)"

    # ---- M3b: Death drains hard, soaks hits, potion dispels ----
    Far; $hp0 = (State).health
    $excel.Run('DebugSpawn', [int]$K_DEATH, [int]58, [int]6)
    $excel.Run('DebugUpdate', [int]210)
    Check "Death drains hard"        ((State).health -le $hp0 - 2*(Dmg $DEATH_DRAIN)) "-$($hp0-(State).health)"
    1..6 | ForEach-Object { $excel.Run('DebugStep', [int]0, [int]0) }
    Check "Death soaks melee hits"   ((State).entN -eq 1) (State).entN
    $excel.Run('DebugSetPotions', [int]1); $excel.Run('DebugUsePotion')
    Check "potion dispels Death"     ((State).entN -eq 0) (State).entN

    # ---- M4a: title / select / characters ----
    $excel.Run('RenderTitle')
    Check "title screen renders"     ((((View-Row 4) -replace ' ','') -like '*GAUNTLEX*')) (View-Row 4).Trim()
    $excel.Run('RenderSelect', [int]$CH_VALKYRIE)
    Check "select screen renders"    ((View-Row 7) -match 'VALKYRIE') (View-Row 7).Trim()

    $excel.Run('DebugSetChar', [int]$CH_ELF); $excel.Run('GameInit')
    Check "character persists init"   ((State).char -eq $CH_ELF) (State).char

    # Valkyrie (armour 60) takes less than Wizard (armour 145) from the same ghost
    $excel.Run('DebugSetChar', [int]$CH_VALKYRIE); Far; $h=(State).health
    $excel.Run('DebugSpawn', [int]$K_GHOST, [int]58, [int]6); $excel.Run('DebugUpdate', [int]120)
    $vLoss = $h - (State).health
    $excel.Run('DebugSetChar', [int]$CH_WIZARD); Far; $h=(State).health
    $excel.Run('DebugSpawn', [int]$K_GHOST, [int]58, [int]6); $excel.Run('DebugUpdate', [int]120)
    $wLoss = $h - (State).health
    Check "armour: valkyrie < wizard"  ($vLoss -lt $wLoss) "valk -$vLoss  wiz -$wLoss"

    # Warrior (hit power 2) destroys a 3-HP generator in 2 melee hits
    $excel.Run('DebugSetChar', [int]$CH_WARRIOR); $excel.Run('GameInit'); $excel.Run('DebugWarp', [int]17, [int]13)
    $b = State; 1..2 | ForEach-Object { $excel.Run('DebugStep', [int]-1, [int]0) }
    Check "warrior 2-hits a generator"  ((State).genAlive -eq $b.genAlive - 1) "gen $($b.genAlive)->$((State).genAlive)"

    # Wizard's potion blast reaches an entity the others cannot
    $excel.Run('DebugSetChar', [int]$CH_WIZARD); $excel.Run('GameInit')
    $excel.Run('DebugSpawn', [int]$K_GRUNT, [int]27, [int]10); $excel.Run('DebugSetPotions', [int]1); $excel.Run('DebugUsePotion')
    $wizKill = ((State).entN -eq 0)
    $excel.Run('DebugSetChar', [int]$CH_VALKYRIE); $excel.Run('GameInit')
    $excel.Run('DebugSpawn', [int]$K_GRUNT, [int]27, [int]10); $excel.Run('DebugSetPotions', [int]1); $excel.Run('DebugUsePotion')
    $valkMiss = ((State).entN -eq 1)
    Check "wizard potion reaches further" ($wizKill -and $valkMiss) "wizKill=$wizKill valkMiss=$valkMiss"

    $excel.Run('DebugSetChar', [int]$CH_WARRIOR)      # restore default for the rest

    # ---- M5a: picture-Shape pool build (headless-safe: no .Fill.UserPicture) ----
    $ok = $true; try { $excel.Run('DebugShapesCompile') } catch { $ok = $false }
    Check "modShapes compiles" $ok $ok
    $excel.Run('GameInit'); $excel.Run('FitViewport'); $excel.Run('RenderInit'); $excel.Run('ShapesInit')
    $mz = 0; $act = 0
    foreach ($s in $ws.Shapes) { if ($s.Name -like 'mz_*') { $mz++ }; if ($s.Name -like 'act_*') { $act++ } }
    Check "shape pools built (240 maze / 80 actor)" ($mz -eq 240 -and $act -eq 80) "mz $mz act $act"

    # ---- M4b: level chaining, victory ----
    $excel.Run('GameInit')
    Check "starts on level 1"        ((State).lvl -eq 1) (State).lvl

    $excel.Run('GameInit')
    $excel.Run('DebugWarp', [int]57, [int]5); $excel.Run('DebugSpawn', [int]$K_GRUNT, [int]57, [int]7)
    $excel.Run('DebugStep', [int]0, [int]1)                    # +10 score
    $excel.Run('DebugSetKeys', [int]2)
    $b = State
    $excel.Run('DebugNextLevel')
    $a = State
    Check "next level: advance + carry" ($a.lvl -eq 2 -and $a.st -eq 'PLAY' -and $a.score -ge $SCORE_GRUNT -and $a.keys -eq 0) "lvl $($a.lvl) $($a.st) score $($a.score) keys $($a.keys)"

    $excel.Run('GameInit'); $excel.Run('DebugSetHealth', [int]500); $excel.Run('DebugNextLevel')
    Check "level bonus tops up health" ((State).health -eq 750) (State).health

    $excel.Run('GameInit')
    $excel.Run('DebugNextLevel'); $excel.Run('DebugNextLevel'); $excel.Run('DebugNextLevel')
    $s = State
    Check "no level after last -> VICTORY" ($s.st -eq 'VICTORY' -and $s.lvl -eq 3) "$($s.st) lvl $($s.lvl)"
    $excel.Run('RenderVictory')
    Check "victory screen renders"    (((View-Row 5) + (View-Row 7)) -match 'ESCAPED|DUNGEON') "$((View-Row 5).Trim())"

    # ---- M2 carry-over ----
    $excel.Run('GameInit'); $excel.Run('DebugWarp', [int]9, [int]43); $excel.Run('DebugStep', [int]0, [int]2)
    Check "key picked up"            ((State).keys -eq 1) (State).keys
    $excel.Run('GameInit'); $excel.Run('DebugWarp', [int]51, [int]57); $excel.Run('DebugStep', [int]2, [int]0)
    Check "door blocks without key"  ((State).plHR -eq 51) (State).plHR
    $excel.Run('DebugWarp', [int]9, [int]43); $excel.Run('DebugStep', [int]0, [int]2)
    $excel.Run('DebugWarp', [int]51, [int]57); $excel.Run('DebugStep', [int]2, [int]0)
    Check "door opens with key"      ((State).plHR -eq 53 -and (State).keys -eq 0) "plHR $((State).plHR)"
    $excel.Run('GameInit'); $excel.Run('DebugWarp', [int]57, [int]57); $excel.Run('DebugUpdate', [int]50)
    Check "exit -> WON +100"         ((State).st -eq 'WON' -and (State).score -eq 100) "$((State).st)"
    $excel.Run('GameInit'); $excel.Run('DebugSetHealth', [int]2); $excel.Run('DebugUpdate', [int]3400)
    $s = State
    Check "death: life lost + respawn" ($s.lives -eq 2 -and $s.health -eq $START_HEALTH -and $s.st -eq 'PLAY') "lives $($s.lives) $($s.st)"
    $excel.Run('DebugSetHealth', [int]1); $excel.Run('DebugUpdate', [int]1200)
    $excel.Run('DebugSetHealth', [int]1); $excel.Run('DebugUpdate', [int]1200)
    Check "game over at 0 lives"     ((State).lives -eq 0 -and (State).st -eq 'OVER') "$((State).lives) $((State).st)"

    if ($fail) { throw "$fail check(s) failed" }
    Write-Host "OK"
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
