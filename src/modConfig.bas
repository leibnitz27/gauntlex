Attribute VB_Name = "modConfig"
Option Explicit

' ============================================================
'  GAUNTLEX - config / constants
' ============================================================

' ---- Viewport (in cells) ----
Public Const VIEW_COLS  As Long = 40
Public Const VIEW_ROWS  As Long = 25
Public Const CELL_PTS   As Double = 18#      ' square cell edge, in points

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
