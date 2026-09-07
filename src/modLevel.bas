Attribute VB_Name = "modLevel"
Option Explicit

' ============================================================
'  Level: load a map from a (very hidden) worksheet into a
'  string grid, and answer collision queries against it.
'
'  gMap(row, col) is 1-based, one character per cell, sized to
'  the MAP (not the viewport). Off-map reads as solid wall.
' ============================================================

Public gMap()   As String
Public gStartR  As Long, gStartC As Long
Public gExitR   As Long, gExitC As Long

Public Sub LoadLevel(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim raw As Variant
    raw = ws.Range(ws.Cells(1, 1), ws.Cells(MAP_ROWS, MAP_COLS)).Value

    ReDim gMap(1 To MAP_ROWS, 1 To MAP_COLS)
    gStartR = 2: gStartC = 2                 ' safe fallback
    gExitR = 0:  gExitC = 0

    Dim r As Long, c As Long, ch As String
    For r = 1 To MAP_ROWS
        For c = 1 To MAP_COLS
            ch = Left$(CStr(raw(r, c)) & " ", 1)
            Select Case ch
                Case T_WALL
                    gMap(r, c) = T_WALL
                Case T_SPAWN
                    gStartR = r: gStartC = c
                    gMap(r, c) = T_FLOOR
                Case T_EXIT
                    gExitR = r: gExitC = c
                    gMap(r, c) = T_EXIT
                Case Else
                    gMap(r, c) = T_FLOOR
            End Select
        Next c
    Next r
End Sub

' Off-grid counts as solid.
Public Function IsWall(ByVal r As Long, ByVal c As Long) As Boolean
    If r < 1 Or r > MAP_ROWS Or c < 1 Or c > MAP_COLS Then
        IsWall = True
    Else
        IsWall = (gMap(r, c) = T_WALL)
    End If
End Function
