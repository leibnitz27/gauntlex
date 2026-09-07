Attribute VB_Name = "modConfig"
Option Explicit

' ============================================================
'  GAUNTLEX - config / constants
' ============================================================

' ---- Viewport: the window of cells we render into ----
'  modRender.FitViewport picks a square cell size that fills the Excel window
'  (one axis exact, the other scrolls), then how many whole cells fit - all
'  capped at the map size. gCellPts / gViewCols / gViewRows hold the result.
Public Const CELL_DEFAULT_PTS As Double = 18#   ' fallback when the window can't be measured
Public Const CELL_MIN_PTS     As Double = 12#
Public Const CELL_MAX_PTS     As Double = 48#
Public Const HUD_PTS          As Double = 64#   ' vertical band reserved below the viewport
Public Const HUD_ROW_PTS      As Double = 16#   ' height of each HUD row
Public Const VIEW_MIN_COLS    As Long = 12
Public Const VIEW_MIN_ROWS    As Long = 8

Public gViewCols As Long
Public gViewRows As Long
Public gCellPts  As Double

' ---- Map: the level grid, always >= the viewport ----
'  1 Excel cell = 1 maze cell (the original's 16px logical cell).
'  Levels may use less but the array/sheet is always MAP_COLS x MAP_ROWS.
Public Const MAP_COLS   As Long = 32
Public Const MAP_ROWS   As Long = 32

Public Const SCREEN_SHEET As String = "Screen"
Public Const LEVEL_SHEET  As String = "L01"

' ---- Timing ----
Public Const TARGET_FPS As Long = 25
Public Const FRAME_MS   As Long = 40         ' ~= 1000 / TARGET_FPS
Public Const MOVE_MS    As Long = 110        ' player advances one cell this often while a key is held

' ---- Tiles ----
Public Const T_FLOOR    As String = "."
Public Const T_WALL     As String = "#"
Public Const T_SPAWN    As String = "S"
Public Const T_EXIT     As String = "X"
Public Const T_PLAYER   As String = "@"

' ---- Colours ----
Public Const CLR_BG     As Long = 0                  ' RGB(0,0,0)
Public Const CLR_FG     As Long = 5308500            ' RGB(84,208,80) - phosphor green
