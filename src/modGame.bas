Attribute VB_Name = "modGame"
Option Explicit

' ============================================================
'  Game state + per-frame update (M3a).
'
'  Entities (grunt / ghost / demon) live in parallel arrays;
'  a K_NONE slot is free. Generators (from the level) spawn
'  their kind on a timer while on screen and take hits to
'  destroy. The player fires in its facing direction; demons
'  fire back. Keys/doors/food/health-drain/death from M2.
' ============================================================

' ---- player ----
Public gPlHR  As Long, gPlHC As Long
Public gFaceDR As Long, gFaceDC As Long
Public gCamR  As Long, gCamC As Long
Public gState As String                  ' "PLAY" | "WON" | "OVER"
Public gHealth As Long, gScore As Long, gKeys As Long, gPotions As Long, gLives As Long

' ---- entities ----
Public gEntN As Long
Public gEntKind() As Long, gEntHR() As Long, gEntHC() As Long, gEntHP() As Long, gEntT() As Long

' ---- generators (index matches modLevel gGen*; 1..gGenN) ----
Public gGenAlive() As Boolean, gGenHP() As Long, gGenT() As Long

' ---- projectiles ----
Public gPrjN As Long
Public gPrjKind() As Long, gPrjHR() As Long, gPrjHC() As Long, gPrjDR() As Long, gPrjDC() As Long

Private mMoveAcc As Long, mDrainAcc As Long, mEntAcc As Long, mPrjAcc As Long, mShotAcc As Long
Private mEntTick As Long

' ============================================================

Public Sub GameInit()
    LoadLevel LEVEL_SHEET
    gPlHR = gStartHR: gPlHC = gStartHC
    gFaceDR = 1: gFaceDC = 0
    gState = "PLAY"
    gHealth = START_HEALTH: gScore = 0: gKeys = 0: gPotions = 0: gLives = START_LIVES

    ReDim gEntKind(1 To MAX_ENT): ReDim gEntHR(1 To MAX_ENT): ReDim gEntHC(1 To MAX_ENT)
    ReDim gEntHP(1 To MAX_ENT): ReDim gEntT(1 To MAX_ENT)
    gEntN = 0

    ReDim gGenAlive(1 To MAX_GEN): ReDim gGenHP(1 To MAX_GEN): ReDim gGenT(1 To MAX_GEN)
    Dim i As Long
    For i = 1 To gGenN
        gGenAlive(i) = True: gGenHP(i) = GEN_HP
        gGenT(i) = 600 + (i * 370) Mod 2000        ' stagger first spawns
    Next i

    ReDim gPrjKind(1 To MAX_PRJ): ReDim gPrjHR(1 To MAX_PRJ): ReDim gPrjHC(1 To MAX_PRJ)
    ReDim gPrjDR(1 To MAX_PRJ): ReDim gPrjDC(1 To MAX_PRJ)
    gPrjN = 0

    mMoveAcc = 0: mDrainAcc = 0: mEntAcc = 0: mPrjAcc = 0: mShotAcc = SHOT_MS
    mEntTick = 0
    CenterCamera
End Sub

Public Sub GameUpdate(ByVal dt As Long)
    If gState <> "PLAY" Then Exit Sub

    mDrainAcc = mDrainAcc + dt
    Do While mDrainAcc >= DRAIN_MS
        mDrainAcc = mDrainAcc - DRAIN_MS: gHealth = gHealth - DRAIN_AMOUNT
    Loop

    mMoveAcc = mMoveAcc + dt
    If mMoveAcc >= MOVE_MS Then mMoveAcc = 0: StepPlayer

    mShotAcc = mShotAcc + dt
    If gInFire And mShotAcc >= SHOT_MS Then mShotAcc = 0: FireShot P_PLAYER, gPlHR + gFaceDR, gPlHC + gFaceDC, gFaceDR, gFaceDC

    mEntAcc = mEntAcc + dt
    Do While mEntAcc >= ENT_TICK_MS
        mEntAcc = mEntAcc - ENT_TICK_MS: mEntTick = mEntTick + 1: StepEntities
    Loop

    RunGenerators dt

    mPrjAcc = mPrjAcc + dt
    Do While mPrjAcc >= PRJ_MS
        mPrjAcc = mPrjAcc - PRJ_MS: StepProjectiles
    Loop

    If gHealth <= 0 Then PlayerDied
    If gState = "PLAY" And gExitBR <> 0 And PlayerCovers(gExitBR, gExitBC) Then
        gScore = gScore + SCORE_EXIT: gState = "WON"
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
                gKeys = gKeys - 1: gBlock(br(i), bc(i)) = T_FLOOR
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
                Case T_FOOD: gHealth = gHealth + FOOD_VALUE: gBlock(br(i), bc(i)) = T_FLOOR
                Case T_KEY:  gKeys = gKeys + 1:              gBlock(br(i), bc(i)) = T_FLOOR
            End Select
        End If
    Next i
End Sub

Private Sub MeleeEntAt(ByVal hr As Long, ByVal hc As Long)
    Dim i As Long
    For i = 1 To gEntN
        If gEntKind(i) <> K_NONE Then
            If Overlap(hr, hc, gEntHR(i), gEntHC(i)) Then KillEntity i
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
    Dim i As Long, k As Long, moveNow As Boolean
    For i = 1 To gEntN
        k = gEntKind(i)
        If k = K_NONE Then GoTo NextE

        moveNow = (k = K_GHOST) Or (mEntTick Mod 2 = 0)
        If moveNow Then MoveToward i, gPlHR, gPlHC

        If k = K_DEMON Then
            gEntT(i) = gEntT(i) - ENT_TICK_MS
            If gEntT(i) <= 0 Then
                gEntT(i) = DEMON_SHOOT_MS
                Dim ddr As Long, ddc As Long
                ddr = Sgn(gPlHR - gEntHR(i)): ddc = Sgn(gPlHC - gEntHC(i))
                If ddr <> 0 Or ddc <> 0 Then FireShot P_ENEMY, gEntHR(i) + ddr, gEntHC(i) + ddc, ddr, ddc
            End If
        End If

        If Overlap(gPlHR, gPlHC, gEntHR(i), gEntHC(i)) Then
            If k = K_GHOST Then
                gHealth = gHealth - GHOST_DMG
                gEntKind(i) = K_NONE               ' kamikaze
            Else
                gHealth = gHealth - GRUNT_TOUCH_DMG
            End If
        End If
NextE:
    Next i
End Sub

Private Sub MoveToward(ByVal i As Long, ByVal tr As Long, ByVal tc As Long)
    Dim dr As Long, dc As Long
    dr = Sgn(tr - gEntHR(i)): dc = Sgn(tc - gEntHC(i))
    If dr = 0 And dc = 0 Then Exit Sub
    If FootprintClear(gEntHR(i) + dr, gEntHC(i) + dc) Then
        gEntHR(i) = gEntHR(i) + dr: gEntHC(i) = gEntHC(i) + dc
    ElseIf dr <> 0 And FootprintClear(gEntHR(i) + dr, gEntHC(i)) Then
        gEntHR(i) = gEntHR(i) + dr
    ElseIf dc <> 0 And FootprintClear(gEntHR(i), gEntHC(i) + dc) Then
        gEntHC(i) = gEntHC(i) + dc
    End If
End Sub

Private Sub KillEntity(ByVal i As Long)
    Select Case gEntKind(i)
        Case K_GRUNT: gScore = gScore + SCORE_GRUNT
        Case K_GHOST: gScore = gScore + SCORE_GHOST
        Case K_DEMON: gScore = gScore + SCORE_DEMON
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
    gEntKind(s) = k: gEntHR(s) = hr: gEntHC(s) = hc: gEntHP(s) = 1
    gEntT(s) = IIf(k = K_DEMON, DEMON_SHOOT_MS, 0)
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
            gGenHP(i) = gGenHP(i) - 1
            If gGenHP(i) <= 0 Then
                gGenAlive(i) = False
                gBlock(br, bc) = T_FLOOR
                gScore = gScore + SCORE_GEN
            End If
            Exit Sub
        End If
    Next i
End Sub

' ---- projectiles ---------------------------------------

Private Sub FireShot(ByVal owner As Long, ByVal hr As Long, ByVal hc As Long, ByVal dr As Long, ByVal dc As Long)
    Dim s As Long, i As Long
    For i = 1 To gPrjN
        If gPrjKind(i) = 0 Then s = i: Exit For
    Next i
    If s = 0 Then
        If gPrjN >= MAX_PRJ Then Exit Sub
        gPrjN = gPrjN + 1: s = gPrjN
    End If
    gPrjKind(s) = owner: gPrjHR(s) = hr: gPrjHC(s) = hc: gPrjDR(s) = dr: gPrjDC(s) = dc
End Sub

Private Sub StepProjectiles()
    Dim i As Long, j As Long
    For i = 1 To gPrjN
        If gPrjKind(i) = 0 Then GoTo NextP
        gPrjHR(i) = gPrjHR(i) + gPrjDR(i)
        gPrjHC(i) = gPrjHC(i) + gPrjDC(i)

        If gPrjHR(i) < 1 Or gPrjHR(i) > MAP_ROWS Or gPrjHC(i) < 1 Or gPrjHC(i) > MAP_COLS Then
            gPrjKind(i) = 0: GoTo NextP
        End If
        If IsWallHC(gPrjHR(i), gPrjHC(i)) Then
            If IsGen(BlockAtHC(gPrjHR(i), gPrjHC(i))) Then
                HitGen (gPrjHR(i) + 1) \ 2, (gPrjHC(i) + 1) \ 2
            End If
            gPrjKind(i) = 0: GoTo NextP
        End If

        If gPrjKind(i) = P_PLAYER Then
            For j = 1 To gEntN
                If gEntKind(j) <> K_NONE Then
                    If PointIn(gPrjHR(i), gPrjHC(i), gEntHR(j), gEntHC(j)) Then
                        KillEntity j: gPrjKind(i) = 0: GoTo NextP
                    End If
                End If
            Next j
        Else
            If PointIn(gPrjHR(i), gPrjHC(i), gPlHR, gPlHC) Then
                gHealth = gHealth - DEMON_SHOT_DMG: gPrjKind(i) = 0: GoTo NextP
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
    Else
        gHealth = START_HEALTH
        gPlHR = gStartHR: gPlHC = gStartHC
        mDrainAcc = 0
        CenterCamera
    End If
End Sub

Private Function InMap(ByVal br As Long, ByVal bc As Long) As Boolean
    InMap = (br >= 1 And br <= MAP_BLOCK_ROWS And bc >= 1 And bc <= MAP_BLOCK_COLS)
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
Public Sub DebugFire(): FireShot P_PLAYER, gPlHR + gFaceDR, gPlHC + gFaceDC, gFaceDR, gFaceDC: End Sub

Public Function DebugState() As String
    Dim i As Long, ea As Long, ga As Long, pa As Long, e1r As Long, e1c As Long, best As Long
    e1r = -1: e1c = -1: best = 1000000
    For i = 1 To gEntN
        If gEntKind(i) <> K_NONE Then
            ea = ea + 1
            Dim d As Long: d = Abs(gEntHR(i) - gPlHR) + Abs(gEntHC(i) - gPlHC)
            If d < best Then best = d: e1r = gEntHR(i): e1c = gEntHC(i)
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
                 ";" & ea & ";" & e1r & ";" & e1c & ";" & ga & ";" & pa
End Function
