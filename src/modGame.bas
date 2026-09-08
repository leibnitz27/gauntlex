Attribute VB_Name = "modGame"
Option Explicit

' ============================================================
'  Game state + per-frame update (M2).
'
'  Player: 2x2 half-cell footprint (gPlHR/gPlHC), one half-cell
'  per MOVE_MS. Health drains every DRAIN_MS. Grunts step toward
'  the player every GRUNT_MS (8-dir greedy + wall-slide) and
'  drain health on contact; walking into a grunt melee-kills it.
'  Keys open doors; food restores health. Out of health -> lose
'  a life and respawn at spawn; out of lives -> game over.
' ============================================================

Public gPlHR  As Long, gPlHC As Long     ' player footprint top-left, half-cell
Public gCamR  As Long, gCamC As Long     ' camera top-left, half-cell
Public gState As String                  ' "PLAY" | "WON" | "OVER"

Public gHealth  As Long
Public gScore   As Long
Public gKeys    As Long
Public gPotions As Long
Public gLives   As Long

Public gGrCount   As Long
Public gGrHR()    As Long, gGrHC() As Long
Public gGrAlive() As Boolean

Private mMoveAcc  As Long
Private mGruntAcc As Long
Private mDrainAcc As Long

Public Sub GameInit()
    LoadLevel LEVEL_SHEET
    gPlHR = gStartHR: gPlHC = gStartHC
    mMoveAcc = 0: mGruntAcc = 0: mDrainAcc = 0
    gState = "PLAY"
    gHealth = START_HEALTH
    gScore = 0
    gKeys = 0
    gPotions = 0
    gLives = START_LIVES
    SpawnGrunts
    CenterCamera
End Sub

Private Sub SpawnGrunts()
    ReDim gGrHR(1 To MAX_GRUNTS): ReDim gGrHC(1 To MAX_GRUNTS): ReDim gGrAlive(1 To MAX_GRUNTS)
    gGrCount = gGruntN
    Dim i As Long
    For i = 1 To gGrCount
        gGrHR(i) = gGruntSR(i): gGrHC(i) = gGruntSC(i): gGrAlive(i) = True
    Next i
End Sub

Public Sub GameUpdate(ByVal dt As Long)
    If gState <> "PLAY" Then Exit Sub

    mDrainAcc = mDrainAcc + dt
    Do While mDrainAcc >= DRAIN_MS
        mDrainAcc = mDrainAcc - DRAIN_MS
        gHealth = gHealth - DRAIN_AMOUNT
    Loop

    mMoveAcc = mMoveAcc + dt
    If mMoveAcc >= MOVE_MS Then mMoveAcc = 0: StepPlayer

    mGruntAcc = mGruntAcc + dt
    If mGruntAcc >= GRUNT_MS Then mGruntAcc = 0: StepGrunts

    If gHealth <= 0 Then PlayerDied
    If gState = "PLAY" Then
        If gExitBR <> 0 And PlayerCovers(gExitBR, gExitBC) Then
            gScore = gScore + SCORE_EXIT
            gState = "WON"
        End If
    End If
End Sub

' ---- player ------------------------------------------------

Private Sub StepPlayer()
    Dim dr As Long, dc As Long
    If gInUp Then dr = -1
    If gInDown Then dr = 1
    If gInLeft Then dc = -1
    If gInRight Then dc = 1
    If dr = 0 And dc = 0 Then Exit Sub

    OpenDoorsAt gPlHR + dr, gPlHC + dc          ' unlock a door directly in the way

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
    MeleeAt gPlHR, gPlHC
    CenterCamera
End Sub

Private Sub OpenDoorsAt(ByVal hr As Long, ByVal hc As Long)
    Dim br() As Long, bc() As Long, n As Long, i As Long
    n = FootprintBlocks(hr, hc, br, bc)
    For i = 1 To n
        If InMap(br(i), bc(i)) Then
            If gBlock(br(i), bc(i)) = T_DOOR And gKeys > 0 Then
                gKeys = gKeys - 1
                gBlock(br(i), bc(i)) = T_FLOOR
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
                Case T_FOOD
                    gHealth = gHealth + FOOD_VALUE
                    gBlock(br(i), bc(i)) = T_FLOOR
                Case T_KEY
                    gKeys = gKeys + 1
                    gBlock(br(i), bc(i)) = T_FLOOR
            End Select
        End If
    Next i
End Sub

Private Sub MeleeAt(ByVal hr As Long, ByVal hc As Long)
    Dim i As Long
    For i = 1 To gGrCount
        If gGrAlive(i) Then
            If Overlap(hr, hc, gGrHR(i), gGrHC(i)) Then
                gGrAlive(i) = False
                gScore = gScore + SCORE_GRUNT
            End If
        End If
    Next i
End Sub

' ---- grunts ----------------------------------------------

Private Sub StepGrunts()
    Dim i As Long, dr As Long, dc As Long
    For i = 1 To gGrCount
        If Not gGrAlive(i) Then GoTo NextG

        dr = Sgn(gPlHR - gGrHR(i))
        dc = Sgn(gPlHC - gGrHC(i))
        If dr <> 0 Or dc <> 0 Then
            If FootprintClear(gGrHR(i) + dr, gGrHC(i) + dc) Then
                gGrHR(i) = gGrHR(i) + dr: gGrHC(i) = gGrHC(i) + dc
            ElseIf dr <> 0 And FootprintClear(gGrHR(i) + dr, gGrHC(i)) Then
                gGrHR(i) = gGrHR(i) + dr
            ElseIf dc <> 0 And FootprintClear(gGrHR(i), gGrHC(i) + dc) Then
                gGrHC(i) = gGrHC(i) + dc
            End If
        End If

        If Overlap(gPlHR, gPlHC, gGrHR(i), gGrHC(i)) Then
            gHealth = gHealth - GRUNT_TOUCH_DMG
        End If
NextG:
    Next i
End Sub

' ---- outcomes -------------------------------------------

Private Sub PlayerDied()
    gLives = gLives - 1
    If gLives <= 0 Then
        gLives = 0
        gState = "OVER"
    Else
        gHealth = START_HEALTH
        gPlHR = gStartHR: gPlHC = gStartHC
        mDrainAcc = 0
        CenterCamera
    End If
End Sub

' ---- helpers -------------------------------------------

Private Function InMap(ByVal br As Long, ByVal bc As Long) As Boolean
    InMap = (br >= 1 And br <= MAP_BLOCK_ROWS And bc >= 1 And bc <= MAP_BLOCK_COLS)
End Function

' two 2x2 half-cell footprints overlap iff their tops are within 1 on each axis
Private Function Overlap(ByVal ar As Long, ByVal ac As Long, _
                         ByVal br As Long, ByVal bc As Long) As Boolean
    Overlap = (Abs(ar - br) <= 1 And Abs(ac - bc) <= 1)
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

' ---- test hooks (build\Smoke-Test.ps1) -----------------

Public Sub DebugStep(ByVal dr As Long, ByVal dc As Long)
    OpenDoorsAt gPlHR + dr, gPlHC + dc
    If FootprintClear(gPlHR + dr, gPlHC + dc) Then
        gPlHR = gPlHR + dr: gPlHC = gPlHC + dc
        PickupAt gPlHR, gPlHC
        MeleeAt gPlHR, gPlHC
    End If
    CenterCamera
End Sub

Public Sub DebugUpdate(ByVal dtMs As Long): GameUpdate dtMs: End Sub
Public Sub DebugSetHealth(ByVal h As Long): gHealth = h: End Sub
Public Sub DebugWarp(ByVal hr As Long, ByVal hc As Long)
    gPlHR = hr: gPlHC = hc: CenterCamera
End Sub

Public Function DebugState() As String
    Dim i As Long, ga As Long, g1r As Long, g1c As Long, best As Long
    g1r = -1: g1c = -1: best = 1000000
    For i = 1 To gGrCount
        If gGrAlive(i) Then
            ga = ga + 1
            Dim d As Long: d = Abs(gGrHR(i) - gPlHR) + Abs(gGrHC(i) - gPlHC)
            If d < best Then best = d: g1r = gGrHR(i): g1c = gGrHC(i)
        End If
    Next i
    DebugState = VIEW_COLS & ";" & VIEW_ROWS & ";" & gCamR & ";" & gCamC & _
                 ";" & gPlHR & ";" & gPlHC & ";" & gState & ";" & Format$(gCellPts, "0.0") & _
                 ";" & gLives & ";" & gHealth & ";" & gKeys & ";" & gScore & _
                 ";" & ga & ";" & g1r & ";" & g1c
End Function
