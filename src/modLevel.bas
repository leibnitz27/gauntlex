Attribute VB_Name = "modLevel"
Option Explicit

' ============================================================
'  Level: load a block-resolution map from a (very hidden)
'  worksheet, and answer collision queries in half-cell space.
'
'  gBlock(row, col) is 1-based, one char per 16px block.
'  Half-cell (hr, hc) maps to block ((hr+1)\2, (hc+1)\2).
'  Items (food/key/door) live in gBlock and are edited in place
'  as the player picks them up / opens them.
' ============================================================

Public gBlock()  As String
Public gStartHR  As Long, gStartHC As Long        ' spawn, half-cell coords
Public gExitBR   As Long, gExitBC As Long         ' exit, block coords
Public gGruntN   As Long                          ' grunt spawn markers found
Public gGruntSR() As Long, gGruntSC() As Long     ' ... their half-cell coords

Public Sub LoadLevel(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim raw As Variant
    raw = ws.Range(ws.Cells(1, 1), ws.Cells(MAP_BLOCK_ROWS, MAP_BLOCK_COLS)).Value

    ReDim gBlock(1 To MAP_BLOCK_ROWS, 1 To MAP_BLOCK_COLS)
    ReDim gGruntSR(1 To MAX_GRUNTS): ReDim gGruntSC(1 To MAX_GRUNTS)
    gStartHR = 3: gStartHC = 3                     ' safe fallback (block 2,2)
    gExitBR = 0:  gExitBC = 0
    gGruntN = 0

    Dim r As Long, c As Long, ch As String
    For r = 1 To MAP_BLOCK_ROWS
        For c = 1 To MAP_BLOCK_COLS
            ch = Left$(CStr(raw(r, c)) & " ", 1)
            Select Case ch
                Case T_WALL, T_FOOD, T_KEY, T_DOOR, T_EXIT
                    gBlock(r, c) = ch
                    If ch = T_EXIT Then gExitBR = r: gExitBC = c
                Case T_SPAWN
                    gStartHR = r * 2 - 1: gStartHC = c * 2 - 1
                    gBlock(r, c) = T_FLOOR
                Case T_GRUNTSP
                    If gGruntN < MAX_GRUNTS Then
                        gGruntN = gGruntN + 1
                        gGruntSR(gGruntN) = r * 2 - 1
                        gGruntSC(gGruntN) = c * 2 - 1
                    End If
                    gBlock(r, c) = T_FLOOR
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

' Solid to movement: walls, closed doors, off-map.
Public Function IsWallHC(ByVal hr As Long, ByVal hc As Long) As Boolean
    Dim b As String
    b = BlockAtHC(hr, hc)
    IsWallHC = (b = T_WALL Or b = T_DOOR)
End Function

' Is the 2x2 footprint with top-left (hr, hc) clear of walls / closed doors?
Public Function FootprintClear(ByVal hr As Long, ByVal hc As Long) As Boolean
    FootprintClear = Not (IsWallHC(hr, hc) Or IsWallHC(hr + 1, hc) _
                       Or IsWallHC(hr, hc + 1) Or IsWallHC(hr + 1, hc + 1))
End Function

' The distinct blocks a 2x2 footprint at (hr, hc) covers (1 or 2 per axis).
' Fills br()/bc() (each ReDim'd 1..n) and returns n (1..4).
Public Function FootprintBlocks(ByVal hr As Long, ByVal hc As Long, _
                                ByRef br() As Long, ByRef bc() As Long) As Long
    Dim rr(1 To 2) As Long, cc(1 To 2) As Long, nr As Long, nc As Long
    rr(1) = (hr + 1) \ 2: nr = 1
    If ((hr + 1 + 1) \ 2) <> rr(1) Then nr = 2: rr(2) = rr(1) + 1
    cc(1) = (hc + 1) \ 2: nc = 1
    If ((hc + 1 + 1) \ 2) <> cc(1) Then nc = 2: cc(2) = cc(1) + 1

    Dim n As Long, i As Long, j As Long
    ReDim br(1 To nr * nc): ReDim bc(1 To nr * nc)
    For i = 1 To nr
        For j = 1 To nc
            n = n + 1: br(n) = rr(i): bc(n) = cc(j)
        Next j
    Next i
    FootprintBlocks = n
End Function
