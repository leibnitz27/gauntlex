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

' ---- Renderer ----
Public Const RENDER_SHAPES_DEFAULT As Boolean = True   ' picture-Shape renderer (M5); False = glyphs
Public gRenderShapes As Boolean
Public Const TILE_WALL  As String = "r12_c28.png"      ' blue solid-wall block
Public Const TILE_FLOOR As String = "r13_c11.png"
Public Const TILE_EXIT  As String = "r12_c34.png"

' ---- Sprite atlas: linear stream, index i -> r{i\44}_c{i mod 44}.png ----
'  walk blocks are frame-major: col_in_block = frame*8 + dir
'  dir order: N NE E SE S SW W NW  (= 0..7)
Public Const SPR_HERO_BASE As Long = 396      ' Warrior; +37 per hero (Valk, Wiz, Elf)
Public Const SPR_HERO_STRIDE As Long = 37     ' 24 walk + 8 weapon + 5 dissolve
Public Const SPR_GHOST     As Long = 0
Public Const SPR_GRUNT     As Long = 24
Public Const SPR_HEAD      As Long = 48       ' "floating head" -> thief
Public Const SPR_IMP       As Long = 72       ' -> demon
Public Const SPR_BADWIZ    As Long = 96       ' -> sorcerer
Public Const SPR_DEATH     As Long = 120
Public Const SPR_BLOB      As Long = 144      ' -> lobber
Public Const SPR_GENERATOR As Long = 224      ' r05 c04 (single tile)
Public Const SPR_FIREBALL  As Long = 323      ' r07 c15, +dir  (demon / enemy shot)
Public Const SPR_ARROW     As Long = 531      ' r12 c03, +dir  (Elf weapon block)
Public Const SPR_ITEM_MEAT As Long = 282
Public Const SPR_ITEM_KEY  As Long = 283
Public Const SPR_ITEM_POT  As Long = 289
Public Const ANIM_MS       As Long = 130      ' walk-frame interval

Public Const SCREEN_SHEET As String = "Screen"

' ---- Progression ----
Public Const LEVEL_CLEAR_BONUS As Long = 250      ' health topped up on entering the next level
Public Const WARN_FOOD_AT      As Long = 700      ' "needs food badly"
Public Const WARN_DIE_AT       As Long = 300      ' "is about to die"
Public Const WARN_REARM_AT     As Long = 950      ' health above this re-arms the warnings

' ---- Timing ----
Public Const TARGET_FPS  As Long = 25
Public Const FRAME_MS    As Long = 40
Public Const MOVE_MS     As Long = 60             ' player: one half-cell step this often while a key is held
Public Const ENT_TICK_MS As Long = 65            ' base monster tick; kinds move every 1-3 ticks
Public Const DRAIN_MS    As Long = 1100           ' health lost per tick just for being alive
Public Const SHOT_MS     As Long = 200            ' player fire cooldown
Public Const PRJ_MS      As Long = 45            ' projectile step interval
Public Const GEN_SPAWN_MS As Long = 900           ' generator spawn interval (while on screen)
Public Const DEMON_SHOOT_MS As Long = 1300
Public Const SORC_FLICKER_MS As Long = 650        ' sorcerer visible<->invisible toggle
Public Const LOBBER_THROW_MS As Long = 1500
Public Const POTION_MS   As Long = 400            ' min gap between potion uses
Public Const THIEF_DELAY_MS As Long = 6000        ' after level start, if the level has a T marker
Public Const DEATH_DELAY_MS As Long = 12000       ' ... a Y marker

' ---- Tiles (block grid) ----
Public Const T_FLOOR    As String = "."
Public Const T_WALL     As String = "#"
Public Const T_SPAWN    As String = "S"
Public Const T_EXIT     As String = "X"
Public Const T_FOOD     As String = "+"          ' restores health
Public Const T_KEY      As String = "K"
Public Const T_DOOR     As String = "D"          ' solid until opened with a key
Public Const T_GEN_GRUNT  As String = "G"        ' generators (each spawns its monster while on screen)
Public Const T_GEN_GHOST  As String = "O"
Public Const T_GEN_DEMON  As String = "E"
Public Const T_GEN_SORC   As String = "Z"
Public Const T_GEN_LOBBER As String = "L"
Public Const T_POTION     As String = "P"        ' pickup
Public Const T_THIEF      As String = "T"        ' spawn markers (cleared to floor at load)
Public Const T_DEATH      As String = "Y"
Public Const T_PLAYER     As String = "@"

' ---- Entity kinds ----
Public Const K_NONE   As Long = 0
Public Const K_GRUNT  As Long = 1
Public Const K_GHOST  As Long = 2
Public Const K_DEMON  As Long = 3
Public Const K_SORC   As Long = 4
Public Const K_LOBBER As Long = 5
Public Const K_THIEF  As Long = 6
Public Const K_DEATH  As Long = 7

Public Const P_PLAYER As Long = 1                ' projectile owner
Public Const P_ENEMY  As Long = 2               ' demon fireball - dies on wall
Public Const P_LOBBER As Long = 3               ' rock - arcs over walls, limited life

' ---- Rules ----
Public Const START_HEALTH    As Long = 2000
Public Const START_LIVES     As Long = 3
Public Const DRAIN_AMOUNT    As Long = 1
Public Const FOOD_VALUE      As Long = 350
Public Const GRUNT_TOUCH_DMG As Long = 3          ' per entity tick while a monster overlaps you
Public Const GHOST_DMG       As Long = 55         ' ghost kamikaze hit (then it dies)
Public Const DEMON_SHOT_DMG  As Long = 70
Public Const LOBBER_DMG      As Long = 50
Public Const LOBBER_RANGE    As Long = 8          ' half-cells: closer than this, it backs off and throws
Public Const LOBBER_ROCK_LIFE As Long = 30       ' rock steps before it lands
Public Const DEMON_SHOT_LIFE As Long = 60
Public Const DEATH_HP        As Long = 120        ' hits to kill Death the hard way
Public Const DEATH_DRAIN     As Long = 7          ' per entity tick while Death overlaps you
Public Const GEN_HP          As Long = 3
Public Const GEN_KIND_CAP    As Long = 12         ' a generator idles if this many of its kind are already alive
Public Const SCORE_GRUNT  As Long = 10
Public Const SCORE_GHOST  As Long = 12
Public Const SCORE_DEMON  As Long = 40
Public Const SCORE_SORC   As Long = 20
Public Const SCORE_LOBBER As Long = 25
Public Const SCORE_THIEF  As Long = 200
Public Const SCORE_DEATH  As Long = 500
Public Const SCORE_GEN    As Long = 100
Public Const SCORE_EXIT   As Long = 100
Public Const MAX_ENT      As Long = 200
Public Const MAX_GEN      As Long = 64
Public Const MAX_PRJ      As Long = 96

' ---- Colours ----
' VBA Const can't call RGB(), so these are precomputed R + G*256 + B*65536.
Public Const CLR_BG     As Long = 0               ' black
Public Const CLR_FG     As Long = 6348880         ' RGB(80,224,96)  - phosphor green
' alt: 43775 = RGB(255,170,0) amber  |  16777215 = white
