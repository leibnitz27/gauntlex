Attribute VB_Name = "modGame"
Option Explicit

' ============================================================
'  Game state + per-frame update.
'  M1: grid movement on a map larger than the viewport, with a
'  follow-camera that recentres on the player and clamps at the
'  level edges. HUD counters are placeholders until M2.
' ============================================================

Public gPlR   As Long, gPlC As Long      ' player row / col   (map coords)
Public gCamR  As Long, gCamC As Long     ' camera top-left    (map coords)
Public gState As String                  ' "PLAY" | "WON"

Public gHealth  As Long                  ' placeholders - drain / pickups arrive in M2
Public gScore   As Long
Public gKeys    As Long
Public gPotions As Long

Private mMoveAcc As Long                  ' ms accumulated toward the next step

Public Sub GameInit()
    LoadLevel LEVEL_SHEET
    gPlR = gStartR
    gPlC = gStartC
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
    If Not IsWall(gPlR + dr, gPlC + dc) Then
        gPlR = gPlR + dr: gPlC = gPlC + dc
    ElseIf dr <> 0 And Not IsWall(gPlR + dr, gPlC) Then
        gPlR = gPlR + dr
    ElseIf dc <> 0 And Not IsWall(gPlR, gPlC + dc) Then
        gPlC = gPlC + dc
    End If

    CenterCamera

    If gExitR <> 0 Then
        If gPlR = gExitR And gPlC = gExitC Then gState = "WON"
    End If
End Sub

' Put the player mid-viewport, then clamp so the window stays on the map.
Public Sub CenterCamera()
    gCamR = gPlR - gViewRows \ 2
    gCamC = gPlC - gViewCols \ 2
    If gCamR < 1 Then gCamR = 1
    If gCamC < 1 Then gCamC = 1
    If gCamR > MAP_ROWS - gViewRows + 1 Then gCamR = MAP_ROWS - gViewRows + 1
    If gCamC > MAP_COLS - gViewCols + 1 Then gCamC = MAP_COLS - gViewCols + 1
End Sub

' Test hook: move the player by (dr, dc) cells, honouring walls at the
' destination, and recentre the camera. Used by build\Smoke-Test.ps1.
Public Sub DebugStep(ByVal dr As Long, ByVal dc As Long)
    If Not IsWall(gPlR + dr, gPlC + dc) Then
        gPlR = gPlR + dr
        gPlC = gPlC + dc
    End If
    CenterCamera
End Sub

' Test hook: state as "viewCols;viewRows;camR;camC;plR;plC;state;cellPts".
Public Function DebugState() As String
    DebugState = gViewCols & ";" & gViewRows & ";" & gCamR & ";" & gCamC & _
                 ";" & gPlR & ";" & gPlC & ";" & gState & _
                 ";" & Format$(gCellPts, "0.0")
End Function
