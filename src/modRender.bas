Attribute VB_Name = "modRender"
Option Explicit

' ============================================================
'  Render: state -> cells. One array blit per frame.
'  Cell interior/font colour is set once in RenderInit, never
'  per frame, so a frame is a single Range.Value assignment.
' ============================================================

Private mView As Range
Private mBuf() As String

Public Sub RenderInit()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SCREEN_SHEET)
    ws.Activate

    Application.ScreenUpdating = False
    ws.Cells.Clear
    With ws.Cells.Font
        .Name = "Consolas"
        .Size = 12
    End With
    ws.Cells.HorizontalAlignment = xlCenter
    ws.Cells.VerticalAlignment = xlCenter

    ws.Rows("1:" & VIEW_ROWS).RowHeight = CELL_PTS
    SquareColumns ws

    Set mView = ws.Range(ws.Cells(1, 1), ws.Cells(VIEW_ROWS, VIEW_COLS))
    mView.Interior.Color = CLR_BG
    mView.Font.Color = CLR_FG
    ReDim mBuf(1 To VIEW_ROWS, 1 To VIEW_COLS)

    ws.Cells(VIEW_ROWS + 2, 1).Font.Bold = True
    Application.ScreenUpdating = True
End Sub

' Excel column width is in "characters"; we want points. Binary-search it.
Private Sub SquareColumns(ByVal ws As Worksheet)
    Dim lo As Double, hi As Double, mid As Double, i As Long
    lo = 0.25: hi = 12#
    For i = 1 To 26
        mid = (lo + hi) / 2
        ws.Columns(1).ColumnWidth = mid
        If ws.Columns(1).Width > CELL_PTS Then hi = mid Else lo = mid
    Next i
    ws.Range(ws.Columns(1), ws.Columns(VIEW_COLS)).ColumnWidth = ws.Columns(1).ColumnWidth
End Sub

Public Sub RenderFrame(ByVal fps As Double)
    Dim r As Long, c As Long
    For r = 1 To VIEW_ROWS
        For c = 1 To VIEW_COLS
            mBuf(r, c) = gMap(r, c)
        Next c
    Next r
    mBuf(gPlR, gPlC) = T_PLAYER

    mView.Value = mBuf

    Dim tail As String
    If gState = "WON" Then
        tail = "   *** LEVEL CLEARED  -  press ESC ***"
    Else
        tail = "   arrows: move    ESC: quit"
    End If
    mView.Worksheet.Cells(VIEW_ROWS + 2, 1).Value = _
        "GAUNTLEX  M0     fps " & Format$(fps, "0") & _
        "     pos " & gPlR & "," & gPlC & tail
End Sub
