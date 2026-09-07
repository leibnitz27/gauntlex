Attribute VB_Name = "modLevel"
Option Explicit

' ============================================================
'  Level: load a block-resolution map from a (very hidden)
'  worksheet, and answer collision queries in half-cell space.
'
'  gBlock(row, col) is 1-based, one char per 16px block.
'  Half-cell (hr, hc) maps to block ((hr+1)\2, (hc+1)\2).
' ============================================================

Public gBlock()  As String
Public gStartHR  As Long, gStartHC As Long        ' spawn, half-cell coords
Public gExitBR   As Long, gExitBC As Long         ' exit, block coords

Public Sub LoadLevel(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim raw As Variant
    raw = ws.Range(ws.Cells(1, 1), ws.Cells(MAP_BLOCK_ROWS, MAP_BLOCK_COLS)).Value

    ReDim gBlock(1 To MAP_BLOCK_ROWS, 1 To MAP_BLOCK_COLS)
    gStartHR = 3: gStartHC = 3                     ' safe fallback (block 2,2)
    gExitBR = 0:  gExitBC = 0

    Dim r As Long, c As Long, ch As String
    For r = 1 To MAP_BLOCK_ROWS
        For c = 1 To MAP_BLOCK_COLS
            ch = Left$(CStr(raw(r, c)) & " ", 1)
            Select Case ch
                Case T_WALL
                    gBlock(r, c) = T_WALL
                Case T_SPAWN
                    gStartHR = r * 2 - 1: gStartHC = c * 2 - 1
                    gBlock(r, c) = T_FLOOR
                Case T_EXIT
                    gExitBR = r: gExitBC = c
                    gBlock(r, c) = T_EXIT
                Case Else
                    gBlock(r, c) = T_FLOOR
            End Select
        Next c
    Next r
End Sub

' Block char at a half-cell (for rendering). Off-map reads as wall.
Public Function BlockAtHC(ByVal hr As Long, ByVal hc As Long) As String
    Dim br As Long, bc As Long
    br = (hr + 1) \ 2
    bc = (hc + 1) \ 2
    If br < 1 Or br > MAP_BLOCK_ROWS Or bc < 1 Or bc > MAP_BLOCK_COLS Then
        BlockAtHC = T_WALL
    Else
        BlockAtHC = gBlock(br, bc)
    End If
End Function

' Is the half-cell inside a wall block? Off-map counts as solid.
Public Function IsWallHC(ByVal hr As Long, ByVal hc As Long) As Boolean
    IsWallHC = (BlockAtHC(hr, hc) = T_WALL)
End Function

' Is the 2x2 footprint with top-left (hr, hc) clear of walls?
Public Function FootprintClear(ByVal hr As Long, ByVal hc As Long) As Boolean
    FootprintClear = Not (IsWallHC(hr, hc) Or IsWallHC(hr + 1, hc) _
                       Or IsWallHC(hr, hc + 1) Or IsWallHC(hr + 1, hc + 1))
End Function
