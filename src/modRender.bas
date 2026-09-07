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

    ' player: a 2x2 half-cell footprint
    Dim dr As Long, dc As Long, vr As Long, vc As Long
    For dr = 0 To 1
        For dc = 0 To 1
            vr = (gPlHR + dr) - gCamR + 1
            vc = (gPlHC + dc) - gCamC + 1
            If vr >= 1 And vr <= VIEW_ROWS And vc >= 1 And vc <= VIEW_COLS Then
                mBuf(vr, vc) = T_PLAYER
            End If
        Next dc
    Next dr

    mView.Value = mBuf
    DrawHud mView.Worksheet, fps
End Sub

Private Sub DrawHud(ByVal ws As Worksheet, ByVal fps As Double)
    Dim h As Long: h = mHudCol
    ws.Cells(2, h).Value = "GAUNTLEX"
    ws.Cells(4, h).Value = "HEALTH   " & Format$(gHealth, "0000")
    ws.Cells(5, h).Value = "SCORE    " & Format$(gScore, "000000")
    ws.Cells(6, h).Value = "KEYS     " & gKeys
    ws.Cells(7, h).Value = "POTIONS  " & gPotions
    ws.Cells(10, h).Value = "fps " & Format$(fps, "0")
    ws.Cells(11, h).Value = "blk " & ((gPlHR + 1) \ 2) & "," & ((gPlHC + 1) \ 2)
    ws.Cells(12, h).Value = "cam " & gCamR & "," & gCamC
    ws.Cells(14, h).Value = IIf(gState = "WON", "*** CLEARED - ESC ***", "arrows move / ESC quit")
End Sub
