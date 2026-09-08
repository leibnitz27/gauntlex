#Requires -Version 5
# Opens the workbook, drives it headless, checks render + M3 game logic.
$ErrorActionPreference = 'Stop'
$book = Join-Path (Split-Path -Parent $PSScriptRoot) 'gauntlex.xlsm'

$VC = 36; $VR = 20
$START_HEALTH = 2000; $START_LIVES = 3
$K_GRUNT=1; $K_GHOST=2; $K_DEMON=3; $K_SORC=4; $K_LOBBER=5; $K_THIEF=6; $K_DEATH=7
$GHOST_DMG=55; $DEMON_SHOT_DMG=70; $LOBBER_DMG=50; $DEATH_DRAIN=7
$SCORE_GRUNT=10; $SCORE_GEN=100; $SCORE_THIEF=200

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
        $first=$null; $n=0
        for ($r=1;$r -le $VR;$r++){for($c=1;$c -le $VC;$c++){
            if ([string]$ws.Cells.Item($r,$c).Value2 -eq $glyph){ if(-not $first){$first="$r,$c"}; $n++ }
        }}
        [pscustomobject]@{ first=$first; n=$n }
    }
    # DebugState -> vc;vr;camR;camC;plHR;plHC;state;cell;lives;health;keys;score;entN;e1r;e1c;genAlive;prjN;potions;e1kind
    function State {
        $p = ([string]$excel.Run('DebugState')).Split(';')
        [pscustomobject]@{
            camR=[int]$p[2]; camC=[int]$p[3]; plHR=[int]$p[4]; plHC=[int]$p[5]; st=$p[6]; cell=[double]$p[7]
            lives=[int]$p[8]; health=[int]$p[9]; keys=[int]$p[10]; score=[int]$p[11]
            entN=[int]$p[12]; e1r=[int]$p[13]; e1c=[int]$p[14]; genAlive=[int]$p[15]; prjN=[int]$p[16]
            pot=[int]$p[17]; e1k=[int]$p[18]
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
    Check "ghost kamikaze"           ((State).entN -eq 0 -and (State).health -eq $hp0 - $GHOST_DMG) "ent $((State).entN) hp -$($hp0-(State).health)"

    Far; $hp0=(State).health; $excel.Run('DebugSpawn', [int]$K_DEMON, [int]57, [int]40); $excel.Run('DebugUpdate', [int]1500)
    Check "demon shot hits player"   ((State).health -le $hp0 - $DEMON_SHOT_DMG) "-$($hp0-(State).health)"

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
    Check "lobber rock crossed wall" ((State).health -le $hp0 - $LOBBER_DMG) "-$($hp0-(State).health)"

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
    Check "Death drains hard"        ((State).health -le $hp0 - 2*$DEATH_DRAIN) "-$($hp0-(State).health)"
    1..6 | ForEach-Object { $excel.Run('DebugStep', [int]0, [int]0) }
    Check "Death soaks melee hits"   ((State).entN -eq 1) (State).entN
    $excel.Run('DebugSetPotions', [int]1); $excel.Run('DebugUsePotion')
    Check "potion dispels Death"     ((State).entN -eq 0) (State).entN

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
