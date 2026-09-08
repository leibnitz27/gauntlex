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
'  The C64 playfield is ~18 x 10 blocks. We show exactly that. FitViewport
'  scales the half-cell size to fill the full window HEIGHT (modern screens are
'  wide, so height is the binding axis); the HUD lives in the spare width to
'  the right of the playfield.
Public Const VIEW_BLOCK_COLS As Long = 18
Public Const VIEW_BLOCK_ROWS As Long = 10
Public Const VIEW_COLS       As Long = VIEW_BLOCK_COLS * 2    ' 36 half-cells
Public Const VIEW_ROWS       As Long = VIEW_BLOCK_ROWS * 2    ' 20 half-cells

Public Const CELL_DEFAULT_PTS As Double = 20#     ' half-cell size when the window can't be measured
Public Const CELL_MIN_PTS     As Double = 8#
Public Const CELL_MAX_PTS     As Double = 60#     ' generous ceiling; real windows land ~24-46

Public Const HUD_PANEL_PTS    As Double = 220#    ' width reserved right of the playfield for the HUD
Public Const HUD_COL_WIDTH    As Double = 24#     ' chars, for the HUD text column
Public Const HUD_GAP_COLS     As Long = 1         ' blank columns between playfield and HUD

Public gCellPts As Double                         ' current half-cell render size, points

Public Const SCREEN_SHEET As String = "Screen"
Public Const LEVEL_SHEET  As String = "L01"

' ---- Timing ----
Public Const TARGET_FPS As Long = 25
Public Const FRAME_MS   As Long = 40
Public Const MOVE_MS    As Long = 60              ' player: one half-cell step this often while a key is held
Public Const GRUNT_MS   As Long = 130             ' grunts: one half-cell step this often
Public Const DRAIN_MS   As Long = 1100            ' health lost per tick just for being alive

' ---- Tiles (block grid) ----
Public Const T_FLOOR    As String = "."
Public Const T_WALL     As String = "#"
Public Const T_SPAWN    As String = "S"
Public Const T_EXIT     As String = "X"
Public Const T_FOOD     As String = "+"          ' restores health
Public Const T_KEY      As String = "K"
Public Const T_DOOR     As String = "D"          ' solid until opened with a key
Public Const T_GRUNTSP  As String = "G"          ' grunt spawn marker (cleared to floor at load)
Public Const T_PLAYER   As String = "@"
Public Const T_GRUNT    As String = "g"

' ---- Rules ----
Public Const START_HEALTH   As Long = 2000
Public Const START_LIVES    As Long = 3
Public Const DRAIN_AMOUNT   As Long = 1
Public Const FOOD_VALUE     As Long = 350
Public Const GRUNT_TOUCH_DMG As Long = 4          ' per grunt tick while a grunt overlaps you
Public Const SCORE_GRUNT    As Long = 10
Public Const SCORE_EXIT     As Long = 100
Public Const MAX_GRUNTS     As Long = 96

' ---- Colours ----
Public Const CLR_BG     As Long = 0
Public Const CLR_FG     As Long = 5308500         ' RGB(84,208,80) - phosphor green
