Attribute VB_Name = "modRender"
Option Explicit

' ============================================================
'  Render: the camera's half-cell window of the map -> the
'  viewport range. One array blit per frame; colours set once
'  in RenderInit.
'
'  The viewport is a fixed VIEW_COLS x VIEW_ROWS half-cells
'  (= VIEW_BLOCK_COLS x VIEW_BLOCK_ROWS blocks). FitViewport
'  only picks the half-cell render size (gCellPts) so the
'  playfield fills the Excel window.
' ============================================================

Private mView As Range
Private mBuf() As String

' Pick a square half-cell size that fills the tighter window axis; the map is
' square so a wide window keeps side gutter. Resize + hit PLAY to refit.
Public Sub FitViewport()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SCREEN_SHEET)
    ws.Activate
    ActiveWindow.Zoom = 100

    Dim uw As Double, uh As Double
    uw = ActiveWindow.UsableWidth
    uh = ActiveWindow.UsableHeight - HUD_PTS
    If uw < 200 Then uw = 200
    If uh < 120 Then uh = 120

    gCellPts = uw / VIEW_COLS
    If uh / VIEW_ROWS < gCellPts Then gCellPts = uh / VIEW_ROWS
    If gCellPts < CELL_MIN_PTS Then gCellPts = CELL_MIN_PTS
    If gCellPts > CELL_MAX_PTS Then gCellPts = CELL_MAX_PTS
End Sub

Public Sub RenderInit()
    If gCellPts <= 0 Then FitViewport
    If gCellPts <= 0 Then gCellPts = CELL_DEFAULT_PTS

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SCREEN_SHEET)
    ws.Activate

    Application.ScreenUpdating = False
    ws.Cells.Clear
    With ws.Cells.Font
        .Name = "Consolas"
        .Size = FontForCell(gCellPts)
    End With
    ws.Cells.HorizontalAlignment = xlCenter
    ws.Cells.VerticalAlignment = xlCenter

    ws.Rows("1:" & (VIEW_ROWS + 6)).RowHeight = gCellPts
    ws.Rows((VIEW_ROWS + 1) & ":" & (VIEW_ROWS + 6)).RowHeight = HUD_ROW_PTS
    SquareColumns ws

    Set mView = ws.Range(ws.Cells(1, 1), ws.Cells(VIEW_ROWS, VIEW_COLS))
    mView.Interior.Color = CLR_BG
    mView.Font.Color = CLR_FG
    ReDim mBuf(1 To VIEW_ROWS, 1 To VIEW_COLS)

    With ws.Range(ws.Cells(VIEW_ROWS + 2, 1), ws.Cells(VIEW_ROWS + 4, 1))
        .Font.Bold = True
        .Font.Size = 11
    End With
    Application.ScreenUpdating = True
End Sub

Private Function FontForCell(ByVal pts As Double) As Double
    FontForCell = Int(pts * 1.1)          ' one glyph spans a 2-cell block
    If FontForCell < 8 Then FontForCell = 8
    If FontForCell > 28 Then FontForCell = 28
End Function

' Excel column width is in "characters"; we want points. Binary-search it.
Private Sub SquareColumns(ByVal ws As Worksheet)
    Dim lo As Double, hi As Double, mid As Double, i As Long
    lo = 0.25: hi = 40#
    For i = 1 To 30
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

    Dim ws As Worksheet: Set ws = mView.Worksheet
    ws.Cells(VIEW_ROWS + 2, 1).Value = _
        "HEALTH " & Format$(gHealth, "0000") & "     SCORE " & Format$(gScore, "000000")
    ws.Cells(VIEW_ROWS + 3, 1).Value = _
        "KEYS " & gKeys & "     POTIONS " & gPotions
    ws.Cells(VIEW_ROWS + 4, 1).Value = _
        "fps " & Format$(fps, "0") & "   blk " & ((gPlHR + 1) \ 2) & "," & ((gPlHC + 1) \ 2) & _
        "   cam " & gCamR & "," & gCamC & _
        IIf(gState = "WON", "   *** LEVEL CLEARED - ESC ***", "   arrows move / ESC quit")
End Sub
