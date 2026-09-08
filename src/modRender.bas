Attribute VB_Name = "modRender"
Option Explicit

' ============================================================
'  Render: the camera's half-cell window of the map -> the
'  viewport range. One array blit per frame; colours set once
'  in RenderInit.
'
'  Viewport is a fixed VIEW_COLS x VIEW_ROWS half-cells
'  (18 x 10 blocks). FitViewport scales the half-cell size to
'  fill the window height; the HUD sits in a text column to the
'  right of the playfield.
' ============================================================

Private mView As Range
Private mBuf() As String
Private mHudCol As Long

' Fill the full window height with the playfield; shrink only if the playfield
' plus the HUD panel would overrun the width. Resize + hit PLAY to refit.
Public Sub FitViewport()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SCREEN_SHEET)
    ws.Activate
    ActiveWindow.Zoom = 100

    Dim uw As Double, uh As Double
    uw = ActiveWindow.UsableWidth
    uh = ActiveWindow.UsableHeight
    If uw < 300 Then uw = 300
    If uh < 160 Then uh = 160

    gCellPts = uh / VIEW_ROWS
    Dim byWidth As Double
    byWidth = (uw - HUD_PANEL_PTS) / VIEW_COLS
    If gCellPts > byWidth Then gCellPts = byWidth
    If gCellPts < CELL_MIN_PTS Then gCellPts = CELL_MIN_PTS
    If gCellPts > CELL_MAX_PTS Then gCellPts = CELL_MAX_PTS
End Sub

Public Sub RenderInit()
    If gCellPts <= 0 Then FitViewport
    If gCellPts <= 0 Then gCellPts = CELL_DEFAULT_PTS

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SCREEN_SHEET)
    ws.Activate
    mHudCol = VIEW_COLS + HUD_GAP_COLS + 1

    Application.ScreenUpdating = False
    ws.Cells.Clear
    ws.Cells.Interior.Color = CLR_BG
    With ws.Cells.Font
        .Name = "Consolas"
        .Color = CLR_FG
        .Size = FontForCell(gCellPts)
    End With
    ws.Cells.HorizontalAlignment = xlCenter
    ws.Cells.VerticalAlignment = xlCenter

    ws.Rows("1:" & VIEW_ROWS).RowHeight = gCellPts
    SquareColumns ws

    Set mView = ws.Range(ws.Cells(1, 1), ws.Cells(VIEW_ROWS, VIEW_COLS))
    ReDim mBuf(1 To VIEW_ROWS, 1 To VIEW_COLS)

    ' HUD text column, left-aligned, its own small font
    ws.Columns(mHudCol).ColumnWidth = HUD_COL_WIDTH
    With ws.Range(ws.Cells(1, mHudCol), ws.Cells(VIEW_ROWS, mHudCol))
        .HorizontalAlignment = xlLeft
        .Font.Size = 13
    End With
    ws.Cells(2, mHudCol).Font.Size = 18
    ws.Cells(2, mHudCol).Font.Bold = True

    Application.ScreenUpdating = True
End Sub

Private Function FontForCell(ByVal pts As Double) As Double
    FontForCell = Int(pts)               ' one glyph per half-cell
    If FontForCell < 8 Then FontForCell = 8
    If FontForCell > 44 Then FontForCell = 44
End Function

' Excel column width is in "characters"; we want points. Binary-search it.
Private Sub SquareColumns(ByVal ws As Worksheet)
    Dim lo As Double, hi As Double, mid As Double, i As Long
    lo = 0.25: hi = 60#
    For i = 1 To 32
        mid = (lo + hi) / 2
        ws.Columns(1).ColumnWidth = mid
        If ws.Columns(1).Width > gCellPts Then hi = mid Else lo = mid
    Next i
    ws.Range(ws.Columns(1), ws.Columns(VIEW_COLS)).ColumnWidth = ws.Columns(1).ColumnWidth
End Sub

Public Sub RenderFrame(ByVal fps As Double)
    Dim r As Long, c As Long
    For r = 1 To VIEW_ROWS
        For c = 1 To VIEW_COLS
            mBuf(r, c) = BlockAtHC(gCamR + r - 1, gCamC + c - 1)
        Next c
    Next r

    ' projectiles (points), entities (2x2), then the player on top
    Dim i As Long
    For i = 1 To gPrjN
        If gPrjKind(i) <> 0 Then Poke gPrjHR(i), gPrjHC(i), IIf(gPrjKind(i) = P_PLAYER, "*", "!")
    Next i
    For i = 1 To gEntN
        Select Case gEntKind(i)
            Case K_GRUNT:  Stamp gEntHR(i), gEntHC(i), "g"
            Case K_GHOST:  Stamp gEntHR(i), gEntHC(i), "o"
            Case K_DEMON:  Stamp gEntHR(i), gEntHC(i), "d"
            Case K_SORC:   If SorcVisible(i) Then Stamp gEntHR(i), gEntHC(i), "z"
            Case K_LOBBER: Stamp gEntHR(i), gEntHC(i), "l"
            Case K_THIEF:  Stamp gEntHR(i), gEntHC(i), "t"
            Case K_DEATH:  Stamp gEntHR(i), gEntHC(i), "Y"
        End Select
    Next i
    Stamp gPlHR, gPlHC, T_PLAYER

    mView.Value = mBuf
    DrawHud mView.Worksheet, fps
End Sub

Private Sub Stamp(ByVal hr As Long, ByVal hc As Long, ByVal glyph As String)
    Dim dr As Long, dc As Long
    For dr = 0 To 1
        For dc = 0 To 1
            Poke hr + dr, hc + dc, glyph
        Next dc
    Next dr
End Sub

Private Sub Poke(ByVal hr As Long, ByVal hc As Long, ByVal glyph As String)
    Dim vr As Long, vc As Long
    vr = hr - gCamR + 1: vc = hc - gCamC + 1
    If vr >= 1 And vr <= VIEW_ROWS And vc >= 1 And vc <= VIEW_COLS Then mBuf(vr, vc) = glyph
End Sub

Private Sub DrawHud(ByVal ws As Worksheet, ByVal fps As Double)
    Dim h As Long: h = mHudCol
    ws.Cells(2, h).Value = "GAUNTLEX"
    ws.Cells(4, h).Value = "HEALTH   " & Format$(gHealth, "0000")
    ws.Cells(5, h).Value = "SCORE    " & Format$(gScore, "000000")
    ws.Cells(6, h).Value = "LIVES    " & gLives
    ws.Cells(7, h).Value = "KEYS     " & gKeys
    ws.Cells(8, h).Value = "POTIONS  " & gPotions
    ws.Cells(10, h).Value = "SPACE fire  C potion"
    ws.Cells(11, h).Value = "fps " & Format$(fps, "0") & "  ent " & CountEnts()
    ws.Cells(12, h).Value = "blk " & ((gPlHR + 1) \ 2) & "," & ((gPlHC + 1) \ 2)
    Dim msg As String
    Select Case gState
        Case "WON":  msg = "*** LEVEL CLEARED - ESC ***"
        Case "OVER": msg = "*** GAME OVER - ESC ***"
        Case Else:   msg = "arrows move / ESC quit"
    End Select
    ws.Cells(14, h).Value = msg
End Sub

Private Function CountEnts() As Long
    Dim i As Long
    For i = 1 To gEntN
        If gEntKind(i) <> K_NONE Then CountEnts = CountEnts + 1
    Next i
End Function
