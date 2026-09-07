Attribute VB_Name = "modRender"
Option Explicit

' ============================================================
'  Render: the camera's window of the map -> the viewport range.
'  One array blit per frame. Cell interior/font colour is set
'  once in RenderInit, never per frame.
'
'  FitViewport picks a square cell size that fills the Excel
'  window (one axis exact, the other scrolls) and how many
'  whole cells fit, capped at the map. RenderInit then builds
'  the range/buffer at gCellPts / gViewCols / gViewRows.
' ============================================================

Private mView As Range
Private mBuf() As String

' Size the viewport + cells to fill the current Excel window (capped at the
' map). The map is square, so a landscape window keeps some side gutter;
' resize the window and hit PLAY again to refit.
Public Sub FitViewport()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SCREEN_SHEET)
    ws.Activate
    ActiveWindow.Zoom = 100

    Dim uw As Double, uh As Double
    uw = ActiveWindow.UsableWidth
    uh = ActiveWindow.UsableHeight - HUD_PTS
    If uw < 200 Then uw = 200
    If uh < 200 Then uh = 200

    ' Cell size that fills the tighter axis against the whole map;
    ' the looser axis then scrolls.
    gCellPts = uw / MAP_COLS
    If uh / MAP_ROWS > gCellPts Then gCellPts = uh / MAP_ROWS
    If gCellPts < CELL_MIN_PTS Then gCellPts = CELL_MIN_PTS
    If gCellPts > CELL_MAX_PTS Then gCellPts = CELL_MAX_PTS

    gViewCols = Int(uw / gCellPts)
    gViewRows = Int(uh / gCellPts)
    If gViewCols > MAP_COLS Then gViewCols = MAP_COLS
    If gViewRows > MAP_ROWS Then gViewRows = MAP_ROWS
    If gViewCols < VIEW_MIN_COLS Then gViewCols = VIEW_MIN_COLS
    If gViewRows < VIEW_MIN_ROWS Then gViewRows = VIEW_MIN_ROWS
End Sub

' Test hook: pin the viewport to a known size at the default cell size.
Public Sub DebugSetViewport(ByVal cols As Long, ByVal rows As Long)
    gViewCols = cols
    gViewRows = rows
    gCellPts = CELL_DEFAULT_PTS
End Sub

Public Sub RenderInit()
    If gViewCols <= 0 Or gViewRows <= 0 Then FitViewport
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

    ws.Rows("1:" & (gViewRows + 6)).RowHeight = gCellPts
    ws.Rows((gViewRows + 1) & ":" & (gViewRows + 6)).RowHeight = HUD_ROW_PTS
    SquareColumns ws

    Set mView = ws.Range(ws.Cells(1, 1), ws.Cells(gViewRows, gViewCols))
    mView.Interior.Color = CLR_BG
    mView.Font.Color = CLR_FG
    ReDim mBuf(1 To gViewRows, 1 To gViewCols)

    With ws.Range(ws.Cells(gViewRows + 2, 1), ws.Cells(gViewRows + 4, 1))
        .Font.Bold = True
        .Font.Size = 11
    End With
    Application.ScreenUpdating = True
End Sub

Private Function FontForCell(ByVal pts As Double) As Double
    FontForCell = Int(pts * 0.62)
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
    ws.Range(ws.Columns(1), ws.Columns(gViewCols)).ColumnWidth = ws.Columns(1).ColumnWidth
End Sub

Public Sub RenderFrame(ByVal fps As Double)
    Dim r As Long, c As Long
    For r = 1 To gViewRows
        For c = 1 To gViewCols
            mBuf(r, c) = gMap(gCamR + r - 1, gCamC + c - 1)
        Next c
    Next r

    Dim pr As Long, pc As Long
    pr = gPlR - gCamR + 1
    pc = gPlC - gCamC + 1
    If pr >= 1 And pr <= gViewRows And pc >= 1 And pc <= gViewCols Then
        mBuf(pr, pc) = T_PLAYER
    End If

    mView.Value = mBuf

    Dim ws As Worksheet: Set ws = mView.Worksheet
    ws.Cells(gViewRows + 2, 1).Value = _
        "HEALTH " & Format$(gHealth, "0000") & "     SCORE " & Format$(gScore, "000000")
    ws.Cells(gViewRows + 3, 1).Value = _
        "KEYS " & gKeys & "     POTIONS " & gPotions
    ws.Cells(gViewRows + 4, 1).Value = _
        "fps " & Format$(fps, "0") & "   pos " & gPlR & "," & gPlC & _
        "   cam " & gCamR & "," & gCamC & "   view " & gViewCols & "x" & gViewRows & _
        IIf(gState = "WON", "   *** LEVEL CLEARED - ESC ***", "   arrows move / ESC quit")
End Sub
