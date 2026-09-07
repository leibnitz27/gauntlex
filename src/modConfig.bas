Attribute VB_Name = "modConfig"
Option Explicit

' ============================================================
'  GAUNTLEX - config / constants
' ============================================================

' ---- Two grids (see README "Three grids") ----
'  BLOCK grid  : 16px logical cell. Walls, doors, items, exits, AI targeting.
'                Level files are authored at this resolution (one char = one block).
'  HALF-CELL   : 8px. 1 Excel cell = 1 half-cell. The player and monsters have a
'                position on this grid and a 2x2 half-cell (= one block) footprint,
'                and move one half-cell at a time.
Public Const MAP_BLOCK_COLS As Long = 32          ' arcade max playfield
Public Const MAP_BLOCK_ROWS As Long = 32
Public Const MAP_COLS       As Long = 64          ' = MAP_BLOCK_COLS * 2 (half-cells)
Public Const MAP_ROWS       As Long = 64

' ---- Viewport ----
'  The C64 playfield is ~18 x 10 blocks. We show exactly that; FitViewport
'  scales the half-cell size to fill the tighter window axis (letterboxing the
'  other - the map is square, a wide window keeps side gutter until we add a
'  side HUD panel).
Public Const VIEW_BLOCK_COLS As Long = 18
Public Const VIEW_BLOCK_ROWS As Long = 10
Public Const VIEW_COLS       As Long = VIEW_BLOCK_COLS * 2    ' 36 half-cells
Public Const VIEW_ROWS       As Long = VIEW_BLOCK_ROWS * 2    ' 20 half-cells

Public Const CELL_DEFAULT_PTS As Double = 16#     ' half-cell size when the window can't be measured
Public Const CELL_MIN_PTS     As Double = 8#
Public Const CELL_MAX_PTS     As Double = 26#     ' half-cell -> block renders up to 52pt
Public Const HUD_PTS          As Double = 64#     ' vertical band reserved below the viewport
Public Const HUD_ROW_PTS      As Double = 16#

Public gCellPts As Double                         ' current half-cell render size, points

Public Const SCREEN_SHEET As String = "Screen"
Public Const LEVEL_SHEET  As String = "L01"

' ---- Timing ----
Public Const TARGET_FPS As Long = 25
Public Const FRAME_MS   As Long = 40
Public Const MOVE_MS    As Long = 60              ' one half-cell step this often while a key is held

' ---- Tiles (block grid) ----
Public Const T_FLOOR    As String = "."
Public Const T_WALL     As String = "#"
Public Const T_SPAWN    As String = "S"
Public Const T_EXIT     As String = "X"
Public Const T_PLAYER   As String = "@"

' ---- Colours ----
Public Const CLR_BG     As Long = 0
Public Const CLR_FG     As Long = 5308500         ' RGB(84,208,80) - phosphor green
