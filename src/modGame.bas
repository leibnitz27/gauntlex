Attribute VB_Name = "modGame"
Option Explicit

' ============================================================
'  Game state + per-frame update.
'
'  The player occupies a 2x2 half-cell footprint with top-left
'  (gPlHR, gPlHC) and moves one half-cell at a time. Walls,
'  the exit and the camera clamp are all in half-cell space
'  (MAP_ROWS/MAP_COLS = 64). HUD counters are placeholders
'  until M2.
' ============================================================

Public gPlHR  As Long, gPlHC As Long     ' player footprint top-left, half-cell
Public gCamR  As Long, gCamC As Long     ' camera top-left, half-cell
Public gState As String                  ' "PLAY" | "WON"

Public gHealth  As Long
Public gScore   As Long
Public gKeys    As Long
Public gPotions As Long

Private mMoveAcc As Long

Public Sub GameInit()
    LoadLevel LEVEL_SHEET
    gPlHR = gStartHR
    gPlHC = gStartHC
    mMoveAcc = 0
    gState = "PLAY"

    gHealth = 2000
    gScore = 0
    gKeys = 0
    gPotions = 0

    CenterCamera
End Sub

Public Sub GameUpdate(ByVal dt As Long)
    If gState <> "PLAY" Then Exit Sub

    mMoveAcc = mMoveAcc + dt
    If mMoveAcc < MOVE_MS Then Exit Sub
    mMoveAcc = 0

    Dim dr As Long, dc As Long
    If gInUp Then dr = -1
    If gInDown Then dr = 1
    If gInLeft Then dc = -1
    If gInRight Then dc = 1
    If dr = 0 And dc = 0 Then Exit Sub

    ' Try the full move; if blocked, slide along whichever axis is clear.
    If FootprintClear(gPlHR + dr, gPlHC + dc) Then
        gPlHR = gPlHR + dr: gPlHC = gPlHC + dc
    ElseIf dr <> 0 And FootprintClear(gPlHR + dr, gPlHC) Then
        gPlHR = gPlHR + dr
    ElseIf dc <> 0 And FootprintClear(gPlHR, gPlHC + dc) Then
        gPlHC = gPlHC + dc
    End If

    CenterCamera
    If PlayerOnExit() Then gState = "WON"
End Sub

' WON when any footprint half-cell sits in the exit block.
Private Function PlayerOnExit() As Boolean
    If gExitBR = 0 Then Exit Function
    Dim dr As Long, dc As Long
    For dr = 0 To 1
        For dc = 0 To 1
            If (gPlHR + dr + 1) \ 2 = gExitBR And (gPlHC + dc + 1) \ 2 = gExitBC Then
                PlayerOnExit = True
                Exit Function
            End If
        Next dc
    Next dr
End Function

' Put the player mid-viewport, then clamp so the window stays on the map.
Public Sub CenterCamera()
    gCamR = gPlHR - VIEW_ROWS \ 2
    gCamC = gPlHC - VIEW_COLS \ 2
    If gCamR < 1 Then gCamR = 1
    If gCamC < 1 Then gCamC = 1
    If gCamR > MAP_ROWS - VIEW_ROWS + 1 Then gCamR = MAP_ROWS - VIEW_ROWS + 1
    If gCamC > MAP_COLS - VIEW_COLS + 1 Then gCamC = MAP_COLS - VIEW_COLS + 1
End Sub

' Test hook: move the player by (dr, dc) half-cells if the footprint is
' clear, and recentre. Used by build\Smoke-Test.ps1.
Public Sub DebugStep(ByVal dr As Long, ByVal dc As Long)
    If FootprintClear(gPlHR + dr, gPlHC + dc) Then
        gPlHR = gPlHR + dr
        gPlHC = gPlHC + dc
    End If
    CenterCamera
End Sub

' Test hook: "viewCols;viewRows;camR;camC;plHR;plHC;state;cellPts".
Public Function DebugState() As String
    DebugState = VIEW_COLS & ";" & VIEW_ROWS & ";" & gCamR & ";" & gCamC & _
                 ";" & gPlHR & ";" & gPlHC & ";" & gState & _
                 ";" & Format$(gCellPts, "0.0")
End Function
