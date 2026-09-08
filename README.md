# GAUNTLEX

The Commodore 64 version of *Gauntlet* (US Gold, 1986), reimplemented as an
Excel spreadsheet driven by VBA. Cells are the tile grid; VBA is the engine.

Not an emulator — the game logic is rebuilt natively in VBA. The **arcade**
original is the reference spec for level data, the enemy roster and AI, and
timing/feel (it is far better documented and fully ripped, and shares the C64's
grid model); the C64 version sets the look and the 512-level structure. Levels
are hand-built for now; an extraction pass against real arcade level data comes
later. See [Reference: the original](#reference-the-original).

## Requirements

- **Windows desktop Excel** (the engine calls `user32` / `kernel32` / `winmm`).
  Mac Excel and Excel on the web will not run it.
- To build from source: **"Trust access to the VBA project object model"**
  (Excel → File → Options → Trust Center → Trust Center Settings → Macro
  Settings). Or run `build\Enable-VBOM.ps1` once.

## Build & run

```powershell
build\Enable-VBOM.ps1        # once, enables module import
build\Build-Workbook.ps1     # produces gauntlex.xlsm from src\ + levels\
```

Open `gauntlex.xlsm`, enable macros, and click **PLAY** on the `Screen` sheet
(or run the `StartGauntlex` macro, Alt+F8).

- **Arrows** — move
- **ESC** — quit the loop

Edited code in the VBA IDE? Run `build\Export-Modules.ps1` to write it back to
`src\` so git sees the change. `src\` is the source of truth; `gauntlex.xlsm`
is a build artifact and is git-ignored.

## Layout

```
src\           VBA modules (text; the real source)
  modConfig    constants: viewport + map size, timing, tiles, colours
  modWin32     API declares - key polling, ms clock, sleep
  modInput     per-frame keyboard poll -> intent flags
  modLevel     load a map sheet -> string grid (map >= viewport); collision
  modGame      game state + per-frame update
  modRender    visible window of the map -> cells (one array blit per frame)
  modEngine    the frame loop
levels\        one *.txt per level; '#' wall  '.' floor  'S' spawn  'X' exit
build\         PowerShell: enable / build / export
```

## How it runs

VBA has no game loop, so `modEngine.StartGauntlex` *is* the loop:
poll → update → render → `DoEvents` → `Sleep` the rest of the frame
(target 25 fps). Input is `GetAsyncKeyState` polled once per frame, not
`Application.OnKey` events — arrows are trapped for the loop's lifetime so
they drive the player rather than the cell cursor. A frame is a single
`Range.Value = array` assignment; colours are set once at init.

## Reference: the original

Findings from MAME's arcade driver, the C64 version, and a faithful browser
reimplementation ([mJastrzebski6/Gauntlet-I-c64][remake]). These drive the
engine design so we don't paint ourselves into a corner.

### Three grids

Every version of Gauntlet — arcade, C64, the remake — uses the same model:

| Grid | Size | What lives on it |
|---|---|---|
| Hardware tile | 8x8 px | background / charset graphics only |
| **Maze cell** (logical) | 16x16 px = 2x2 tiles | walls, doors, generators, keys, potions, exits; **AI decides on this grid** |
| Movement sub-cell | 8x8 px = 1/2 maze cell | entity position + collision; each actor fills a 2x2 sub-cell footprint |

The sub-cell grid is what lets an actor stand "half in" a doorway. The remake
stores its map at maze-cell resolution and doubles it to the sub-cell grid at
load; monsters step 8 px, decisions step 16 px.

### Numbers

|  | Arcade (1985) | C64 (Gremlin / US Gold '86) | Remake (reference impl) |
|---|---|---|---|
| Playfield tilemap | **64x64 tiles @ 8px** (512x512 px) | — | maps are **33x33 maze cells** |
| Max maze grid | **32x32 cells** (64 / 2) | same model | 33x33 |
| Screen | 336x240 | 320x200 (40x25 char) | 1285x960 canvas |
| HUD | alpha layer, overlaid top + bottom | bottom bar | bottom bar, ~21% of height |
| Visible maze cells | **~20x11** (HUD eats top + bottom) | not authoritatively sourced | ~16x10 (renders 17x11 for scroll) |
| Camera | smooth scroll, map > screen | scrolls | smooth pixel-scroll, follows player, **clamps at edges** |
| Levels | 100+ unique, then cycles | **512** (many are variants) | 8 shipped |
| Players | 4 | **2** | 1 |

**Confidence.** The 64x64 @ 8px playfield and 336x240 screen are from MAME's
[`gauntlet.cpp`][mame] — solid. The 16 px maze cell / 32x32 grid is corroborated
by [StrategyWiki][sw] ("32x32 format... base 16 grid, nothing drawn less than
16x16") and the remake's structure. The C64's exact play-window and HUD height
are **not** authoritatively sourced (the [c64-wiki page][c64w] confirms 512
levels + 2 players but nothing on layout); the remake is the best proxy.

### Enemy AI (from the remake, as a starting spec)

8-direction greedy pursuit toward the nearest player by squared distance, actors
processed nearest-first, on a fixed tick (not every frame). Death absorbs ~170
hits then dies; the sorcerer flickers invisible; the demon fires projectiles.
Generators spawn an actor into a free adjacent cell on a per-generator timer.

### What GAUNTLEX takes from this

- **Two grids, both used.** Levels are authored at **block** resolution (16 px:
  `gBlock`, up to 32x32, `levels/*.txt` one char per block) - walls, doors,
  items, exits, AI targeting live here. The player and monsters live on the
  **half-cell** grid (8 px, up to 64x64): **1 Excel cell = 1 half-cell**, an
  actor is a 2x2 half-cell footprint, and it moves one half-cell at a time.
  Half-cell moves are what make the C64 feel smooth - the player glides rather
  than lurching a whole block, and the camera scrolls in half-cells too.
- **Map is always larger than the viewport.** Camera follows the player (in
  half-cells) and clamps at level edges.
- **Viewport: a fixed 18x10 blocks** (`VIEW_BLOCK_COLS/ROWS`) = 36x20
  half-cells, matching the C64 playfield. `modRender.FitViewport` scales the
  half-cell render size (`gCellPts`) so the playfield fills the full window
  **height** - modern screens are wide, so height is the binding axis and the
  spare width goes to the HUD. It only shrinks below that if the playfield +
  HUD panel would overrun the width. Resize + PLAY to refit.
- **HUD: a text column to the right of the playfield** (`HUD_PANEL_PTS` of
  reserved width) - GAUNTLEX / health / score / keys / potions, arcade-style.

[remake]: https://github.com/mJastrzebski6/Gauntlet-I-c64
[mame]: https://github.com/mamedev/mame/blob/master/src/mame/atari/gauntlet.cpp
[sw]: https://strategywiki.org/wiki/Gauntlet
[c64w]: https://www.c64-wiki.com/wiki/Gauntlet

## Milestones

- [x] **M0** — square-cell viewport, frame loop + fps, arrow movement, wall
      slide, one hand-drawn maze with spawn + exit. *(Viewport == map, no
      camera — the simplification M1 removes.)*

- [x] **M1 — camera & world.** The two-grid world model (see above): block
      levels, half-cell actors, follow-camera in half-cells clamping at edges,
      windowed blit, HUD rows. Viewport is a fixed 18x10 blocks; `FitViewport`
      scales the half-cell size to the window. Verified by
      `build\Smoke-Test.ps1` (12 checks: fit sizing, 2x2 wall/player patches,
      camera scroll + clamp, view-relative player).
  - `modConfig`: `MAP_BLOCK_*` 32x32, `MAP_*` 64x64 half-cells, `VIEW_*`
    fixed at 36x20 half-cells.
  - `modLevel`: `gBlock` (32x32); `BlockAtHC` / `IsWallHC` / `FootprintClear`
    answer in half-cell space. `levels/L01.txt` is a 32x32 block maze.
  - `modGame`: player footprint top-left `gPlHR/gPlHC`, one-half-cell steps
    with 2x2 collision + wall-slide; `CenterCamera` clamps to
    `[1, MAP_ROWS - VIEW_ROWS + 1]`.
  - `modRender`: `FitViewport` picks `gCellPts` to fill window height;
    `RenderFrame` blits `BlockAtHC(gCamR+r-1, gCamC+c-1)` (walls/exit render as
    2x2 quads), stamps the player's 2x2 footprint, and `DrawHud` writes the
    right-hand text column (health / score / keys / potions placeholders).
  - *History: the first M1 pass used a single 16px grid with whole-cell moves;
    reworked to the half-cell model to match the C64's sub-block movement.*

- [x] **M2 — first real level.** `levels/L01.txt` redrawn with keys, a locked
      door guarding the exit, food, and 7 grunt spawns. `modLevel` parses the
      new tiles (`+ K D G`), tracks grunt spawn points, treats closed doors as
      solid. `modGame`: health drains every `DRAIN_MS`; grunts step toward the
      player every `GRUNT_MS` (8-dir greedy + wall-slide) and drain health on
      contact; walking into a grunt melee-kills it; keys open doors; food heals;
      0 health = lose a life + respawn, 0 lives = `OVER`. `modRender` stamps
      grunts (`g`) under the player; HUD gains `LIVES`. Verified by
      `build\Smoke-Test.ps1` (20 checks). *Deferred: generators (M3), wall-
      following AI, ranged attack, tuning.*

- [x] **M3a — entities, generators, fire.** Typed entity arrays
      (`gEnt{Kind,HR,HC,HP,T}`, `K_NONE` = free slot) replace the M2 grunt
      arrays. Level tiles `G/O/E` are grunt/ghost/demon **generators** - solid,
      3 HP, spawn their kind into a free neighbour every `GEN_SPAWN_MS` while
      on screen (capped per kind), destroyed by melee/shots (+100). Player
      **fires** (`SPACE`) in its facing direction; projectiles (`gPrj*`) hit
      walls/entities/generators; **demons fire back**. Ghost is a kamikaze
      (big hit, then gone). `build\Check-Level.ps1` checks generators are
      approachable. Verified by `build\Smoke-Test.ps1` (25 checks).
- [x] **M3b — specials & potions.** Sorcerer (`Z` gen; flickers invisible on a
      derived phase, still solid/dangerous), lobber (`L` gen; backs off inside
      `LOBBER_RANGE` and throws `P_LOBBER` rocks that ignore walls, limited
      life). Potions: `P` pickup, `C` uses one - screen-clear blast kills every
      entity + generator + enemy shot in view (Death included). Thief (`T`
      marker, arrives after `THIEF_DELAY_MS`): steals a key/potion on contact
      then flees fast; kill it to get the item back (+200). Death (`Y` marker,
      after `DEATH_DELAY_MS`): slow, `DEATH_DRAIN`/tick on contact,
      `DEATH_HP` hits to kill the hard way, or one potion. `DamageEntity`
      routes melee/shots (Death soaks, everything else dies at once). Verified
      by `build\Smoke-Test.ps1` (28 checks).

- [ ] **M4 — game shape.** Four characters + select screen, level chaining +
      loader, title screen, sound.

- [ ] **M5 — fidelity pass.** Authentic level data, AI/timing tuned to the
      original, 2-player co-op, and the **picture-Shape renderer** (decided by
      the `spike-renderer` spike): the glyph blit is replaced by two Shape
      pools — a fixed ~180-Shape maze grid (never moved; camera scroll changes
      each Shape's tile) and a ~64-Shape actor pool (moved per frame, tile
      swapped only on facing/anim change), textured from `reference/genesis-tiles/`.
      Coloured-cell rendering was measured slower and dropped.
