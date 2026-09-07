Attribute VB_Name = "modGame"
Option Explicit

' ============================================================
'  Game state + per-frame update.
'  M0: one player token, grid movement, wall slide, reach exit.
' ============================================================

Public gPlR   As Long, gPlC As Long      ' player row / col
Public gState As String                  ' "PLAY" | "WON"

Private mMoveAcc As Long                  ' ms accumulated toward the next step

Public Sub GameInit()
    LoadLevel LEVEL_SHEET
    gPlR = gStartR
    gPlC = gStartC
    mMoveAcc = 0
    gState = "PLAY"
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

    If gExitR <> 0 Then
        If gPlR = gExitR And gPlC = gExitC Then gState = "WON"
    End If
End Sub
