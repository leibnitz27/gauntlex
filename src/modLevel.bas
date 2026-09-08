Attribute VB_Name = "modLevel"
Option Explicit

' ============================================================
'  Level: load a block-resolution map from a (very hidden)
'  worksheet, and answer collision queries in half-cell space.
'
'  gBlock(row, col) is 1-based, one char per 16px block.
'  Half-cell (hr, hc) maps to block ((hr+1)\2, (hc+1)\2).
'  Items (food/key/door) and generators live in gBlock and are
'  edited in place as the game changes them.
' ============================================================

Public gBlock()  As String
Public gStartHR  As Long, gStartHC As Long        ' spawn, half-cell coords
Public gExitBR   As Long, gExitBC As Long         ' exit, block coords

Public gGenN     As Long                          ' generators found in the level
Public gGenBR()  As Long, gGenBC() As Long        ' ... block coords
Public gGenKind() As Long                         ' ... K_GRUNT / K_GHOST / K_DEMON

Public Sub LoadLevel(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim raw As Variant
    raw = ws.Range(ws.Cells(1, 1), ws.Cells(MAP_BLOCK_ROWS, MAP_BLOCK_COLS)).Value

    ReDim gBlock(1 To MAP_BLOCK_ROWS, 1 To MAP_BLOCK_COLS)
    ReDim gGenBR(1 To MAX_GEN): ReDim gGenBC(1 To MAX_GEN): ReDim gGenKind(1 To MAX_GEN)
    gStartHR = 3: gStartHC = 3
    gExitBR = 0:  gExitBC = 0
    gGenN = 0

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
                Case T_GEN_GRUNT, T_GEN_GHOST, T_GEN_DEMON
                    gBlock(r, c) = ch
                    If gGenN < MAX_GEN Then
                        gGenN = gGenN + 1
                        gGenBR(gGenN) = r: gGenBC(gGenN) = c
                        Select Case ch
                            Case T_GEN_GRUNT: gGenKind(gGenN) = K_GRUNT
                            Case T_GEN_GHOST: gGenKind(gGenN) = K_GHOST
                            Case T_GEN_DEMON: gGenKind(gGenN) = K_DEMON
                        End Select
                    End If
                Case Else
                    gBlock(r, c) = T_FLOOR
            End Select
        Next c
    Next r
End Sub

Public Function IsGen(ByVal ch As String) As Boolean
    IsGen = (ch = T_GEN_GRUNT Or ch = T_GEN_GHOST Or ch = T_GEN_DEMON)
End Function

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

' Solid to movement: walls, closed doors, live generators, off-map.
Public Function IsWallHC(ByVal hr As Long, ByVal hc As Long) As Boolean
    Dim b As String
    b = BlockAtHC(hr, hc)
    IsWallHC = (b = T_WALL Or b = T_DOOR Or IsGen(b))
End Function

Public Function FootprintClear(ByVal hr As Long, ByVal hc As Long) As Boolean
    FootprintClear = Not (IsWallHC(hr, hc) Or IsWallHC(hr + 1, hc) _
                       Or IsWallHC(hr, hc + 1) Or IsWallHC(hr + 1, hc + 1))
End Function

' The distinct blocks a 2x2 footprint at (hr, hc) covers (1-4).
Public Function FootprintBlocks(ByVal hr As Long, ByVal hc As Long, _
                                ByRef br() As Long, ByRef bc() As Long) As Long
    Dim rr(1 To 2) As Long, cc(1 To 2) As Long, nr As Long, nc As Long
    rr(1) = (hr + 1) \ 2: nr = 1
    If ((hr + 2) \ 2) <> rr(1) Then nr = 2: rr(2) = rr(1) + 1
    cc(1) = (hc + 1) \ 2: nc = 1
    If ((hc + 2) \ 2) <> cc(1) Then nc = 2: cc(2) = cc(1) + 1

    Dim n As Long, i As Long, j As Long
    ReDim br(1 To nr * nc): ReDim bc(1 To nr * nc)
    For i = 1 To nr
        For j = 1 To nc
            n = n + 1: br(n) = rr(i): bc(n) = cc(j)
        Next j
    Next i
    FootprintBlocks = n
End Function
