Attribute VB_Name = "modGame"
Option Explicit

' ============================================================
'  Game state + per-frame update (M3b).
'
'  Entities in parallel arrays; K_NONE slot = free. gEntT is
'  a per-kind scratch field: demon/lobber throw cooldown,
'  thief carried-item code (0 none / 1 key / 2 potion). Death
'  uses gEntHP (many hits). Sorcerer visibility is derived
'  from mEntTick + slot, no state.
'  Generators spawn their kind on screen; potions (C) clear
'  the view; the thief steals; Death only a potion stops.
' ============================================================

' ---- player ----
Public gPlHR  As Long, gPlHC As Long
Public gFaceDR As Long, gFaceDC As Long
Public gCamR  As Long, gCamC As Long
Public gChar  As Long                    ' CHAR_WARRIOR..CHAR_ELF
Public gState As String                  ' "TITLE" | "SELECT" | "PLAY" | "WON" | "OVER"
Public gHealth As Long, gScore As Long, gKeys As Long, gPotions As Long, gLives As Long

' ---- entities ----
Public gEntN As Long
Public gEntKind() As Long, gEntHR() As Long, gEntHC() As Long, gEntHP() As Long, gEntT() As Long
Public gEntDir() As Long                  ' 0..7 facing (N NE E SE S SW W NW) - for the sprite renderer

' ---- generators (index matches modLevel gGen*; 1..gGenN) ----
Public gGenAlive() As Boolean, gGenHP() As Long, gGenT() As Long

' ---- projectiles ----
Public gPrjN As Long
Public gPrjKind() As Long, gPrjHR() As Long, gPrjHC() As Long
Public gPrjDR() As Long, gPrjDC() As Long, gPrjLife() As Long

Public gLevelNum As Long
Public gBlocksChanged As Boolean          ' a gBlock cell was edited this frame (shape renderer re-textures)

Private mMoveAcc As Long, mDrainAcc As Long, mEntAcc As Long, mPrjAcc As Long, mShotAcc As Long
Private mPotionAcc As Long, mEntTick As Long
Private mThiefT As Long, mDeathT As Long           ' countdowns; -1 = no marker in level
Private mWarnFood As Boolean, mWarnDie As Boolean

' ============================================================

' full reset - new run from level 1 (also the smoke-test entry point)
Public Sub GameInit()
    gScore = 0: gLives = START_LIVES: gPotions = 0
    gHealth = START_HEALTH
    gLevelNum = 1
    LoadCurrentLevel
End Sub

Private Function LevelName(ByVal n As Long) As String
    LevelName = "L" & Format$(n, "00")
End Function

Private Function LevelSheetExists(ByVal n As Long) As Boolean
    Dim s As Worksheet, nm As String
    nm = LevelName(n)
    For Each s In ThisWorkbook.Worksheets
        If s.Name = nm Then LevelSheetExists = True: Exit Function
    Next s
End Function

' load gLevelNum; keep score / lives / potions / health / char
Private Sub LoadCurrentLevel()
    LoadLevel LevelName(gLevelNum)
    gPlHR = gStartHR: gPlHC = gStartHC
    gFaceDR = 1: gFaceDC = 0
    gState = "PLAY"
    gKeys = 0

    ReDim gEntKind(1 To MAX_ENT): ReDim gEntHR(1 To MAX_ENT): ReDim gEntHC(1 To MAX_ENT)
    ReDim gEntHP(1 To MAX_ENT): ReDim gEntT(1 To MAX_ENT): ReDim gEntDir(1 To MAX_ENT)
    gEntN = 0

    ReDim gGenAlive(1 To MAX_GEN): ReDim gGenHP(1 To MAX_GEN): ReDim gGenT(1 To MAX_GEN)
    Dim i As Long
    For i = 1 To gGenN
        gGenAlive(i) = True: gGenHP(i) = GEN_HP
        gGenT(i) = 600 + (i * 370) Mod 2000
    Next i

    ReDim gPrjKind(1 To MAX_PRJ): ReDim gPrjHR(1 To MAX_PRJ): ReDim gPrjHC(1 To MAX_PRJ)
    ReDim gPrjDR(1 To MAX_PRJ): ReDim gPrjDC(1 To MAX_PRJ): ReDim gPrjLife(1 To MAX_PRJ)
    gPrjN = 0

    mMoveAcc = 0: mDrainAcc = 0: mEntAcc = 0: mPrjAcc = 0: mShotAcc = SHOT_MS
    mPotionAcc = POTION_MS: mEntTick = 0
    mWarnFood = False: mWarnDie = False
    mThiefT = IIf(gThiefBR > 0, THIEF_DELAY_MS, -1)
    mDeathT = IIf(gDeathBR > 0, DEATH_DELAY_MS, -1)
    gBlocksChanged = True            ' shape renderer: re-texture the whole maze
    CenterCamera
    Say "Enter level " & gLevelNum
End Sub

' called by the engine after the level-clear pause; False -> no more levels
Public Function AdvanceLevel() As Boolean
    If Not LevelSheetExists(gLevelNum + 1) Then Exit Function
    gLevelNum = gLevelNum + 1
    gHealth = gHealth + LEVEL_CLEAR_BONUS
    If gHealth > START_HEALTH Then gHealth = START_HEALTH
    LoadCurrentLevel
    AdvanceLevel = True
End Function

Public Sub GameUpdate(ByVal dt As Long)
    If gState <> "PLAY" Then Exit Sub

    mDrainAcc = mDrainAcc + dt
    Do While mDrainAcc >= DRAIN_MS
        mDrainAcc = mDrainAcc - DRAIN_MS: gHealth = gHealth - DRAIN_AMOUNT
    Loop

    mMoveAcc = mMoveAcc + dt
    If mMoveAcc >= CharMoveMs(gChar) Then mMoveAcc = 0: StepPlayer

    mShotAcc = mShotAcc + dt
    If gInFire And mShotAcc >= CharShotMs(gChar) Then
        mShotAcc = 0
        FireShot P_PLAYER, gPlHR + gFaceDR, gPlHC + gFaceDC, gFaceDR, gFaceDC, DEMON_SHOT_LIFE
    End If

    mPotionAcc = mPotionAcc + dt
    If gInPotion And gPotions > 0 And mPotionAcc >= POTION_MS Then mPotionAcc = 0: UsePotion

    ExpireTimers dt

    mEntAcc = mEntAcc + dt
    Do While mEntAcc >= ENT_TICK_MS
        mEntAcc = mEntAcc - ENT_TICK_MS: mEntTick = mEntTick + 1: StepEntities
    Loop

    RunGenerators dt

    mPrjAcc = mPrjAcc + dt
    Do While mPrjAcc >= PRJ_MS
        mPrjAcc = mPrjAcc - PRJ_MS: StepProjectiles
    Loop

    HealthWarnings

    If gHealth <= 0 Then PlayerDied
    If gState = "PLAY" And gExitBR <> 0 And PlayerCovers(gExitBR, gExitBC) Then
        gScore = gScore + SCORE_EXIT: gState = "WON"
        Say "Level complete"
    End If
End Sub

Private Sub HealthWarnings()
    If gHealth > WARN_REARM_AT Then mWarnFood = False: mWarnDie = False
    If gHealth <= WARN_DIE_AT And Not mWarnDie Then
        mWarnDie = True: Say CharName(gChar) & " is about to die"
    ElseIf gHealth <= WARN_FOOD_AT And Not mWarnFood Then
        mWarnFood = True: Say CharName(gChar) & " needs food badly"
    End If
End Sub

' ---- player -------------------------------------------------

Private Sub StepPlayer()
    Dim dr As Long, dc As Long
    If gInUp Then dr = -1
    If gInDown Then dr = 1
    If gInLeft Then dc = -1
    If gInRight Then dc = 1
    If dr <> 0 Or dc <> 0 Then gFaceDR = dr: gFaceDC = dc
    If dr = 0 And dc = 0 Then Exit Sub

    OpenDoorsAt gPlHR + dr, gPlHC + dc
    MeleeGenAt gPlHR + dr, gPlHC + dc          ' bumping a generator hits it

    Dim moved As Boolean
    If FootprintClear(gPlHR + dr, gPlHC + dc) Then
        gPlHR = gPlHR + dr: gPlHC = gPlHC + dc: moved = True
    ElseIf dr <> 0 And FootprintClear(gPlHR + dr, gPlHC) Then
        gPlHR = gPlHR + dr: moved = True
    ElseIf dc <> 0 And FootprintClear(gPlHR, gPlHC + dc) Then
        gPlHC = gPlHC + dc: moved = True
    End If
    If Not moved Then Exit Sub

    PickupAt gPlHR, gPlHC
    MeleeEntAt gPlHR, gPlHC
    CenterCamera
End Sub

Private Sub OpenDoorsAt(ByVal hr As Long, ByVal hc As Long)
    Dim br() As Long, bc() As Long, n As Long, i As Long
    n = FootprintBlocks(hr, hc, br, bc)
    For i = 1 To n
        If InMap(br(i), bc(i)) Then
            If gBlock(br(i), bc(i)) = T_DOOR And gKeys > 0 Then
                gKeys = gKeys - 1: gBlock(br(i), bc(i)) = T_FLOOR: gBlocksChanged = True
            End If
        End If
    Next i
End Sub

Private Sub PickupAt(ByVal hr As Long, ByVal hc As Long)
    Dim br() As Long, bc() As Long, n As Long, i As Long
    n = FootprintBlocks(hr, hc, br, bc)
    For i = 1 To n
        If InMap(br(i), bc(i)) Then
            Select Case gBlock(br(i), bc(i))
                Case T_FOOD:   gHealth = gHealth + FOOD_VALUE: gBlock(br(i), bc(i)) = T_FLOOR: gBlocksChanged = True
                Case T_KEY:    gKeys = gKeys + 1:              gBlock(br(i), bc(i)) = T_FLOOR: gBlocksChanged = True
                Case T_POTION: gPotions = gPotions + 1:        gBlock(br(i), bc(i)) = T_FLOOR: gBlocksChanged = True
            End Select
        End If
    Next i
End Sub

Private Sub MeleeEntAt(ByVal hr As Long, ByVal hc As Long)
    Dim i As Long
    For i = 1 To gEntN
        If gEntKind(i) <> K_NONE Then
            If Overlap(hr, hc, gEntHR(i), gEntHC(i)) Then DamageEntity i
        End If
    Next i
End Sub

Private Sub MeleeGenAt(ByVal hr As Long, ByVal hc As Long)
    Dim br() As Long, bc() As Long, n As Long, i As Long
    n = FootprintBlocks(hr, hc, br, bc)
    For i = 1 To n
        If InMap(br(i), bc(i)) And IsGen(gBlock(br(i), bc(i))) Then HitGen br(i), bc(i)
    Next i
End Sub

' ---- entities --------------------------------------------

Private Sub StepEntities()
    Dim i As Long, k As Long
    For i = 1 To gEntN
        k = gEntKind(i)
        Select Case k
            Case K_NONE
            Case K_GRUNT:  StepChaser i, K_GRUNT, GRUNT_TOUCH_DMG
            Case K_GHOST:  StepGhost i
            Case K_DEMON:  StepDemon i
            Case K_SORC:   StepChaser i, K_SORC, GRUNT_TOUCH_DMG
            Case K_LOBBER: StepLobber i
            Case K_THIEF:  StepThief i
            Case K_DEATH:  StepDeath i
        End Select
    Next i
End Sub

Private Sub Hurt(ByVal amount As Long)               ' combat damage, scaled by armour
    gHealth = gHealth - (amount * CharArmourPct(gChar) \ 100)
End Sub

Private Sub StepChaser(ByVal i As Long, ByVal k As Long, ByVal dmg As Long)
    If mEntTick Mod 2 = 0 Then MoveToward i, gPlHR, gPlHC
    If Overlap(gPlHR, gPlHC, gEntHR(i), gEntHC(i)) Then Hurt dmg
End Sub

Private Sub StepGhost(ByVal i As Long)
    MoveToward i, gPlHR, gPlHC                        ' every tick - fast
    If Overlap(gPlHR, gPlHC, gEntHR(i), gEntHC(i)) Then
        Hurt GHOST_DMG
        gEntKind(i) = K_NONE                          ' kamikaze
    End If
End Sub

Private Sub StepDemon(ByVal i As Long)
    If mEntTick Mod 2 = 0 Then MoveToward i, gPlHR, gPlHC
    gEntT(i) = gEntT(i) - ENT_TICK_MS
    If gEntT(i) <= 0 Then
        gEntT(i) = DEMON_SHOOT_MS
        Dim dr As Long, dc As Long
        dr = Sgn(gPlHR - gEntHR(i)): dc = Sgn(gPlHC - gEntHC(i))
        If dr <> 0 Or dc <> 0 Then FireShot P_ENEMY, gEntHR(i) + dr, gEntHC(i) + dc, dr, dc, DEMON_SHOT_LIFE
    End If
    If Overlap(gPlHR, gPlHC, gEntHR(i), gEntHC(i)) Then Hurt GRUNT_TOUCH_DMG
End Sub

Private Sub StepLobber(ByVal i As Long)
    Dim d As Long: d = Abs(gPlHR - gEntHR(i)) + Abs(gPlHC - gEntHC(i))
    If mEntTick Mod 2 = 0 Then
        If d > LOBBER_RANGE Then
            MoveToward i, gPlHR, gPlHC
        Else
            MoveToward i, gEntHR(i) - Sgn(gPlHR - gEntHR(i)), gEntHC(i) - Sgn(gPlHC - gEntHC(i))  ' back away
        End If
    End If
    gEntT(i) = gEntT(i) - ENT_TICK_MS
    If gEntT(i) <= 0 Then
        gEntT(i) = LOBBER_THROW_MS
        Dim dr As Long, dc As Long
        dr = Sgn(gPlHR - gEntHR(i)): dc = Sgn(gPlHC - gEntHC(i))
        If dr <> 0 Or dc <> 0 Then
            Dim s As Long: s = FireShot(P_LOBBER, gEntHR(i) + dr, gEntHC(i) + dc, dr, dc, LOBBER_ROCK_LIFE)
        End If
    End If
    If Overlap(gPlHR, gPlHC, gEntHR(i), gEntHC(i)) Then Hurt GRUNT_TOUCH_DMG
End Sub

Private Sub StepThief(ByVal i As Long)
    ' fast; chases until it has stolen, then flees
    Dim carrying As Boolean: carrying = (gEntT(i) <> 0)
    If carrying Then
        MoveToward i, gEntHR(i) - Sgn(gPlHR - gEntHR(i)) * 3, gEntHC(i) - Sgn(gPlHC - gEntHC(i)) * 3
        MoveToward i, gEntHR(i) - Sgn(gPlHR - gEntHR(i)) * 3, gEntHC(i) - Sgn(gPlHC - gEntHC(i)) * 3
    Else
        MoveToward i, gPlHR, gPlHC
        If Overlap(gPlHR, gPlHC, gEntHR(i), gEntHC(i)) Then
            If gKeys > 0 Then
                gKeys = gKeys - 1: gEntT(i) = 1
            ElseIf gPotions > 0 Then
                gPotions = gPotions - 1: gEntT(i) = 2
            End If
        End If
    End If
End Sub

Private Sub StepDeath(ByVal i As Long)
    If mEntTick Mod 3 = 0 Then MoveToward i, gPlHR, gPlHC     ' slow
    If Overlap(gPlHR, gPlHC, gEntHR(i), gEntHC(i)) Then Hurt DEATH_DRAIN
End Sub

' sorcerer flicker - render mirrors this
Public Function SorcVisible(ByVal i As Long) As Boolean
    SorcVisible = (((mEntTick \ 10) + i) Mod 2 = 0)
End Function

Private Sub MoveToward(ByVal i As Long, ByVal tr As Long, ByVal tc As Long)
    Dim dr As Long, dc As Long
    dr = Sgn(tr - gEntHR(i)): dc = Sgn(tc - gEntHC(i))
    If dr = 0 And dc = 0 Then Exit Sub
    gEntDir(i) = DirIndex(dr, dc)
    If FootprintClear(gEntHR(i) + dr, gEntHC(i) + dc) Then
        gEntHR(i) = gEntHR(i) + dr: gEntHC(i) = gEntHC(i) + dc
    ElseIf dr <> 0 And FootprintClear(gEntHR(i) + dr, gEntHC(i)) Then
        gEntHR(i) = gEntHR(i) + dr
    ElseIf dc <> 0 And FootprintClear(gEntHR(i), gEntHC(i) + dc) Then
        gEntHC(i) = gEntHC(i) + dc
    End If
End Sub

' one melee/shot hit - Death soaks many, everything else dies at once
Private Sub DamageEntity(ByVal i As Long)
    If gEntKind(i) = K_DEATH Then
        gEntHP(i) = gEntHP(i) - CharHitPower(gChar)
        If gEntHP(i) <= 0 Then KillEntity i
    Else
        KillEntity i
    End If
End Sub

Private Sub KillEntity(ByVal i As Long)
    Select Case gEntKind(i)
        Case K_GRUNT:  gScore = gScore + SCORE_GRUNT
        Case K_GHOST:  gScore = gScore + SCORE_GHOST
        Case K_DEMON:  gScore = gScore + SCORE_DEMON
        Case K_SORC:   gScore = gScore + SCORE_SORC
        Case K_LOBBER: gScore = gScore + SCORE_LOBBER
        Case K_DEATH:  gScore = gScore + SCORE_DEATH
        Case K_THIEF
            gScore = gScore + SCORE_THIEF
            If gEntT(i) = 1 Then gKeys = gKeys + 1        ' drop what it stole
            If gEntT(i) = 2 Then gPotions = gPotions + 1
    End Select
    gEntKind(i) = K_NONE
End Sub

Private Function FreeEntSlot() As Long
    Dim i As Long
    For i = 1 To gEntN
        If gEntKind(i) = K_NONE Then FreeEntSlot = i: Exit Function
    Next i
    If gEntN < MAX_ENT Then gEntN = gEntN + 1: FreeEntSlot = gEntN Else FreeEntSlot = 0
End Function

Private Function CountKind(ByVal k As Long) As Long
    Dim i As Long
    For i = 1 To gEntN
        If gEntKind(i) = k Then CountKind = CountKind + 1
    Next i
End Function

Private Sub SpawnEntity(ByVal k As Long, ByVal hr As Long, ByVal hc As Long)
    Dim s As Long: s = FreeEntSlot()
    If s = 0 Then Exit Sub
    gEntKind(s) = k: gEntHR(s) = hr: gEntHC(s) = hc
    gEntHP(s) = IIf(k = K_DEATH, DEATH_HP, 1)
    Select Case k
        Case K_DEMON:  gEntT(s) = DEMON_SHOOT_MS
        Case K_LOBBER: gEntT(s) = LOBBER_THROW_MS
        Case Else:     gEntT(s) = 0
    End Select
End Sub

' ---- generators -----------------------------------------

Private Sub RunGenerators(ByVal dt As Long)
    Dim i As Long
    For i = 1 To gGenN
        If Not gGenAlive(i) Then GoTo NextG
        gGenT(i) = gGenT(i) - dt
        If gGenT(i) > 0 Then GoTo NextG
        gGenT(i) = GEN_SPAWN_MS

        Dim ghr As Long, ghc As Long
        ghr = gGenBR(i) * 2 - 1: ghc = gGenBC(i) * 2 - 1
        If Not NearCamera(ghr, ghc, 6) Then GoTo NextG
        If CountKind(gGenKind(i)) >= GEN_KIND_CAP Then GoTo NextG

        Dim dr As Long, dc As Long, br As Long, bc As Long
        For dr = -1 To 1
            For dc = -1 To 1
                If dr <> 0 Or dc <> 0 Then
                    br = gGenBR(i) + dr: bc = gGenBC(i) + dc
                    If InMap(br, bc) Then
                        Dim shr As Long, shc As Long
                        shr = br * 2 - 1: shc = bc * 2 - 1
                        If FootprintClear(shr, shc) Then
                            SpawnEntity gGenKind(i), shr, shc
                            GoTo NextG
                        End If
                    End If
                End If
            Next dc
        Next dr
NextG:
    Next i
End Sub

Private Sub HitGen(ByVal br As Long, ByVal bc As Long)
    Dim i As Long
    For i = 1 To gGenN
        If gGenAlive(i) And gGenBR(i) = br And gGenBC(i) = bc Then
            gGenHP(i) = gGenHP(i) - CharHitPower(gChar)
            If gGenHP(i) <= 0 Then
                gGenAlive(i) = False
                gBlock(br, bc) = T_FLOOR: gBlocksChanged = True
                gScore = gScore + SCORE_GEN
            End If
            Exit Sub
        End If
    Next i
End Sub

' ---- projectiles ---------------------------------------

Private Function FireShot(ByVal owner As Long, ByVal hr As Long, ByVal hc As Long, _
                          ByVal dr As Long, ByVal dc As Long, ByVal life As Long) As Long
    Dim s As Long, i As Long
    For i = 1 To gPrjN
        If gPrjKind(i) = 0 Then s = i: Exit For
    Next i
    If s = 0 Then
        If gPrjN >= MAX_PRJ Then Exit Function
        gPrjN = gPrjN + 1: s = gPrjN
    End If
    gPrjKind(s) = owner: gPrjHR(s) = hr: gPrjHC(s) = hc
    gPrjDR(s) = dr: gPrjDC(s) = dc: gPrjLife(s) = life
    FireShot = s
End Function

Private Sub StepProjectiles()
    Dim i As Long, j As Long
    For i = 1 To gPrjN
        If gPrjKind(i) = 0 Then GoTo NextP
        gPrjHR(i) = gPrjHR(i) + gPrjDR(i)
        gPrjHC(i) = gPrjHC(i) + gPrjDC(i)
        gPrjLife(i) = gPrjLife(i) - 1
        If gPrjLife(i) <= 0 Then gPrjKind(i) = 0: GoTo NextP

        If gPrjHR(i) < 1 Or gPrjHR(i) > MAP_ROWS Or gPrjHC(i) < 1 Or gPrjHC(i) > MAP_COLS Then
            gPrjKind(i) = 0: GoTo NextP
        End If
        If gPrjKind(i) <> P_LOBBER And IsWallHC(gPrjHR(i), gPrjHC(i)) Then
            If IsGen(BlockAtHC(gPrjHR(i), gPrjHC(i))) Then
                HitGen (gPrjHR(i) + 1) \ 2, (gPrjHC(i) + 1) \ 2
            End If
            gPrjKind(i) = 0: GoTo NextP
        End If

        If gPrjKind(i) = P_PLAYER Then
            For j = 1 To gEntN
                If gEntKind(j) <> K_NONE Then
                    If PointIn(gPrjHR(i), gPrjHC(i), gEntHR(j), gEntHC(j)) Then
                        DamageEntity j: gPrjKind(i) = 0: GoTo NextP
                    End If
                End If
            Next j
        Else
            If PointIn(gPrjHR(i), gPrjHC(i), gPlHR, gPlHC) Then
                Hurt IIf(gPrjKind(i) = P_LOBBER, LOBBER_DMG, DEMON_SHOT_DMG)
                gPrjKind(i) = 0: GoTo NextP
            End If
        End If
NextP:
    Next i
End Sub

' ---- outcomes / helpers -------------------------------

Private Sub PlayerDied()
    gLives = gLives - 1
    If gLives <= 0 Then
        gLives = 0: gState = "OVER"
        Say "Game over"
    Else
        gHealth = START_HEALTH
        gPlHR = gStartHR: gPlHC = gStartHC
        mDrainAcc = 0
        mWarnFood = False: mWarnDie = False
        CenterCamera
        Say CharName(gChar) & " has died"
    End If
End Sub

' the thief and Death arrive a while into the level (if the level places them)
Private Sub ExpireTimers(ByVal dt As Long)
    If mThiefT >= 0 Then
        mThiefT = mThiefT - dt
        If mThiefT <= 0 Then
            mThiefT = -1
            SpawnEntity K_THIEF, gThiefBR * 2 - 1, gThiefBC * 2 - 1
        End If
    End If
    If mDeathT >= 0 Then
        mDeathT = mDeathT - dt
        If mDeathT <= 0 Then
            mDeathT = -1
            SpawnEntity K_DEATH, gDeathBR * 2 - 1, gDeathBC * 2 - 1
        End If
    End If
End Sub

' screen-clear blast: everything visible dies (Death included)
Private Sub UsePotion()
    gPotions = gPotions - 1
    Dim i As Long, m As Long: m = CharPotionMargin(gChar)
    For i = 1 To gEntN
        If gEntKind(i) <> K_NONE And NearCamera(gEntHR(i), gEntHC(i), m) Then KillEntity i
    Next i
    For i = 1 To gGenN
        If gGenAlive(i) And NearCamera(gGenBR(i) * 2 - 1, gGenBC(i) * 2 - 1, m) Then
            gGenAlive(i) = False
            gBlock(gGenBR(i), gGenBC(i)) = T_FLOOR: gBlocksChanged = True
            gScore = gScore + SCORE_GEN
        End If
    Next i
    For i = 1 To gPrjN
        If gPrjKind(i) <> 0 And gPrjKind(i) <> P_PLAYER Then
            If NearCamera(gPrjHR(i), gPrjHC(i), m) Then gPrjKind(i) = 0
        End If
    Next i
End Sub

Private Function InMap(ByVal br As Long, ByVal bc As Long) As Boolean
    InMap = (br >= 1 And br <= MAP_BLOCK_ROWS And bc >= 1 And bc <= MAP_BLOCK_COLS)
End Function

' (dr,dc) in {-1,0,1} -> facing 0..7 = N NE E SE S SW W NW  (0,0 -> S)
Public Function DirIndex(ByVal dr As Long, ByVal dc As Long) As Long
    Dim t As Variant: t = Array(7, 0, 1, 6, 4, 2, 5, 4, 3)
    DirIndex = t((dr + 1) * 3 + (dc + 1))
End Function

Private Function Overlap(ByVal ar As Long, ByVal ac As Long, ByVal br As Long, ByVal bc As Long) As Boolean
    Overlap = (Abs(ar - br) <= 1 And Abs(ac - bc) <= 1)
End Function

Private Function PointIn(ByVal pr As Long, ByVal pc As Long, ByVal fr As Long, ByVal fc As Long) As Boolean
    PointIn = (pr >= fr And pr <= fr + 1 And pc >= fc And pc <= fc + 1)
End Function

Private Function NearCamera(ByVal hr As Long, ByVal hc As Long, ByVal margin As Long) As Boolean
    NearCamera = (hr >= gCamR - margin And hr <= gCamR + VIEW_ROWS + margin _
             And hc >= gCamC - margin And hc <= gCamC + VIEW_COLS + margin)
End Function

Private Function PlayerCovers(ByVal tbr As Long, ByVal tbc As Long) As Boolean
    Dim br() As Long, bc() As Long, n As Long, i As Long
    n = FootprintBlocks(gPlHR, gPlHC, br, bc)
    For i = 1 To n
        If br(i) = tbr And bc(i) = tbc Then PlayerCovers = True: Exit Function
    Next i
End Function

Public Sub CenterCamera()
    gCamR = gPlHR - VIEW_ROWS \ 2
    gCamC = gPlHC - VIEW_COLS \ 2
    If gCamR < 1 Then gCamR = 1
    If gCamC < 1 Then gCamC = 1
    If gCamR > MAP_ROWS - VIEW_ROWS + 1 Then gCamR = MAP_ROWS - VIEW_ROWS + 1
    If gCamC > MAP_COLS - VIEW_COLS + 1 Then gCamC = MAP_COLS - VIEW_COLS + 1
End Sub

' ---- test hooks --------------------------------------

Public Sub DebugStep(ByVal dr As Long, ByVal dc As Long)
    If dr <> 0 Or dc <> 0 Then gFaceDR = dr: gFaceDC = dc
    OpenDoorsAt gPlHR + dr, gPlHC + dc
    MeleeGenAt gPlHR + dr, gPlHC + dc
    If FootprintClear(gPlHR + dr, gPlHC + dc) Then
        gPlHR = gPlHR + dr: gPlHC = gPlHC + dc
        PickupAt gPlHR, gPlHC
        MeleeEntAt gPlHR, gPlHC
    End If
    CenterCamera
End Sub

Public Sub DebugUpdate(ByVal dtMs As Long): GameUpdate dtMs: End Sub
Public Sub DebugSetHealth(ByVal h As Long): gHealth = h: End Sub
Public Sub DebugWarp(ByVal hr As Long, ByVal hc As Long)
    gPlHR = hr: gPlHC = hc: CenterCamera
End Sub
Public Sub DebugSpawn(ByVal k As Long, ByVal hr As Long, ByVal hc As Long): SpawnEntity k, hr, hc: End Sub
Public Sub DebugFace(ByVal dr As Long, ByVal dc As Long): gFaceDR = dr: gFaceDC = dc: End Sub
Public Sub DebugFire()
    FireShot P_PLAYER, gPlHR + gFaceDR, gPlHC + gFaceDC, gFaceDR, gFaceDC, DEMON_SHOT_LIFE
End Sub
Public Sub DebugSetKeys(ByVal k As Long): gKeys = k: End Sub
Public Sub DebugSetPotions(ByVal p As Long): gPotions = p: End Sub
Public Sub DebugSetChar(ByVal c As Long): gChar = c: End Sub
Public Sub DebugSetRenderShapes(ByVal enabled As Boolean): gRenderShapes = enabled: End Sub
Public Sub DebugNextLevel()              ' mirrors the engine's post-WON step
    If AdvanceLevel() Then gState = "PLAY" Else gState = "VICTORY"
End Sub
Public Sub DebugKillGens()                       ' test isolation - stop all generators
    Dim i As Long
    For i = 1 To gGenN: gGenAlive(i) = False: Next i
End Sub
Public Sub DebugUsePotion(): If gPotions > 0 Then UsePotion
End Sub
Public Function DebugSorcVisible() As Boolean
    Dim i As Long
    For i = 1 To gEntN
        If gEntKind(i) = K_SORC Then DebugSorcVisible = SorcVisible(i): Exit Function
    Next i
End Function

Public Function DebugState() As String
    Dim i As Long, ea As Long, ga As Long, pa As Long
    Dim e1r As Long, e1c As Long, e1k As Long, best As Long
    e1r = -1: e1c = -1: e1k = 0: best = 1000000
    For i = 1 To gEntN
        If gEntKind(i) <> K_NONE Then
            ea = ea + 1
            Dim d As Long: d = Abs(gEntHR(i) - gPlHR) + Abs(gEntHC(i) - gPlHC)
            If d < best Then best = d: e1r = gEntHR(i): e1c = gEntHC(i): e1k = gEntKind(i)
        End If
    Next i
    For i = 1 To gGenN
        If gGenAlive(i) Then ga = ga + 1
    Next i
    For i = 1 To gPrjN
        If gPrjKind(i) <> 0 Then pa = pa + 1
    Next i
    DebugState = VIEW_COLS & ";" & VIEW_ROWS & ";" & gCamR & ";" & gCamC & _
                 ";" & gPlHR & ";" & gPlHC & ";" & gState & ";" & Format$(gCellPts, "0.0") & _
                 ";" & gLives & ";" & gHealth & ";" & gKeys & ";" & gScore & _
                 ";" & ea & ";" & e1r & ";" & e1c & ";" & ga & ";" & pa & _
                 ";" & gPotions & ";" & e1k & ";" & gChar & ";" & gLevelNum
End Function
