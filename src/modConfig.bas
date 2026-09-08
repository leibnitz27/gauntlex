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
Public Const TARGET_FPS  As Long = 25
Public Const FRAME_MS    As Long = 40
Public Const MOVE_MS     As Long = 60             ' player: one half-cell step this often while a key is held
Public Const ENT_TICK_MS As Long = 65            ' base monster tick; kinds move every 1-3 ticks
Public Const DRAIN_MS    As Long = 1100           ' health lost per tick just for being alive
Public Const SHOT_MS     As Long = 200            ' player fire cooldown
Public Const PRJ_MS      As Long = 45            ' projectile step interval
Public Const GEN_SPAWN_MS As Long = 2600          ' generator spawn interval (while on screen)
Public Const DEMON_SHOOT_MS As Long = 1300

' ---- Tiles (block grid) ----
Public Const T_FLOOR    As String = "."
Public Const T_WALL     As String = "#"
Public Const T_SPAWN    As String = "S"
Public Const T_EXIT     As String = "X"
Public Const T_FOOD     As String = "+"          ' restores health
Public Const T_KEY      As String = "K"
Public Const T_DOOR     As String = "D"          ' solid until opened with a key
Public Const T_GEN_GRUNT As String = "G"         ' generators (each spawns its monster while on screen)
Public Const T_GEN_GHOST As String = "O"
Public Const T_GEN_DEMON As String = "E"
Public Const T_PLAYER   As String = "@"

' ---- Entity kinds ----
Public Const K_NONE  As Long = 0
Public Const K_GRUNT As Long = 1
Public Const K_GHOST As Long = 2
Public Const K_DEMON As Long = 3

Public Const P_PLAYER As Long = 1                ' projectile owner
Public Const P_ENEMY  As Long = 2

' ---- Rules ----
Public Const START_HEALTH    As Long = 2000
Public Const START_LIVES     As Long = 3
Public Const DRAIN_AMOUNT    As Long = 1
Public Const FOOD_VALUE      As Long = 350
Public Const GRUNT_TOUCH_DMG As Long = 3          ' per entity tick while a monster overlaps you
Public Const GHOST_DMG       As Long = 55         ' ghost kamikaze hit (then it dies)
Public Const DEMON_SHOT_DMG  As Long = 70
Public Const GEN_HP          As Long = 3
Public Const GEN_KIND_CAP    As Long = 5          ' a generator idles if this many of its kind are already alive
Public Const SCORE_GRUNT As Long = 10
Public Const SCORE_GHOST As Long = 12
Public Const SCORE_DEMON As Long = 40
Public Const SCORE_GEN   As Long = 100
Public Const SCORE_EXIT  As Long = 100
Public Const MAX_ENT     As Long = 200
Public Const MAX_GEN     As Long = 64
Public Const MAX_PRJ     As Long = 80

' ---- Colours ----
Public Const CLR_BG     As Long = 0
Public Const CLR_FG     As Long = 5308500         ' RGB(84,208,80) - phosphor green
