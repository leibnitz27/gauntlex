# Reference material

Visual reference for the port. Not used by the build.

## C64 Gauntlet (Gremlin / US Gold) — the target look & feel

| file | shows |
|---|---|
| `gauntlet-c64-012.webp` | room with a ghost swarm; Valkyrie + key + potion; the ~16px block grid |
| `gauntlet-c64-022.webp` | open area, scattered ghosts + treasure chests |
| `gauntlet-c64-034.webp` | EXIT block, L-shaped brick wall — clearest wall-tile read |
| `gauntlet-c64-044.webp` | corridor with mummies/ghosts, ice-brick walls |

Establishes: ~18×10 block playfield, chunky HUD, smooth sub-block (half-cell)
movement, black surround.

## Genesis / Mega Drive Gauntlet IV (1993, M2/Tengen) — the tile skin for M5

`genesis-gauntlet-samples.png` — 795×256, magenta (`#FF00FF`) transparency key.
A **44 × 14 grid of 16×16 tiles**, pitch 18 (16px tile + 2px grey separator),
origin (3, 5).

`slice-tiles.ps1` cuts it into `genesis-tiles/r{RR}_c{CC}.png` (16×16, magenta
→ alpha, non-empty cells only — 592 tiles). `genesis-tiles-contact.png` is a
3× labelled contact sheet.

Row contents (each row is an 8-direction walk cycle unless noted). Rough until
confirmed tile-by-tile against `tilesheet.png` (`build\Make-TileSheet.ps1`).

| rows | contents |
|---|---|
| 00 | white ghosts, then brown/horned demons |
| 01 | red/orange grunts, small imps |
| 02 | red-white robed sorcerers, then blue-robed enemy wizards |
| 03 | enemy wizards, green blobs, orange ring-burst (spawn/magic fx) |
| 04 | blue armoured figure, then red Valkyrie (shield + sword) |
| 05 | generator box, explosion/scatter frames, yellow magic-sparkle, projectiles |
| 06 | EXIT / "EXIT TO n" signs, floor panels, keys, blue/orange potion bottles, teleporter |
| 07 | floor panels, ankh, score badge, chests, shot frames, score pop-ups (1000..8000), fire |
| 08 | large red dragon/boss frames, then brown-haired Elf |
| 09 | Elf walk, then thrown daggers/shuriken |
| 10 | blonde Warrior walk, then thrown swords/crosses |
| 11 | Warrior frames, flame frames, then green goblins (lobbers) |
| 12 | **Elf** (c00-15: c00-02 frames, c03-10 arrow x8 dirs, c11-15 portal-shrink), then walls |
| 13 | thin blue wall pieces (c00-10), then floor variants |

### Tile map (confirmed with the user, in progress)

| tile(s) | is |
|---|---|
| `r12 c00-c02` | Elf hero frames |
| `r12 c03-c10` | Elf's arrow (fired), 8 directions |
| `r12 c11-c15` | Elf shrinking into a portal (teleport) |
| `r12 c23`/`c24`/`c25`/`c26`/`c28`/`c29`/`c30`/`c31` | **solid wall block**, one per level palette (brown / orange / yellow / black / blue / red / green / white). Same 3x3-bar hash shape, no orientation. |
| `r13 c25`/`c26` | green solid wall block (as above) |
| `r12 c34`, `r12 c35` | **EXIT** block |
| `r13 c11-c24` | brown **floor** variants (visual noise) |
| `r13 c27-c43` | green / blue-tinted floor variants |

### Sprite stream (confirmed with the user)

The atlas is a **linear stream**: stream index `i` -> tile `r{i\44}_c{i mod 44}`
(only non-empty tiles were sliced). Walk blocks are **frame-major**:
`col_in_block = frame*8 + dir`, `dir` = {N,NE,E,SE,S,SW,W,NW} = 0..7.

Heroes - contiguous from r09 c00, each a 37-tile block
(24 walk [frame*8+dir, 3 frames] + 8 thrown-weapon [1/dir] + 5 dissolve-to-exit):

| hero | stream start | = tile |
|---|---|---|
| Warrior (barbarian) | 396 | r09 c00 |
| Valkyrie | 433 | r09 c37 |
| Wizard | 470 | r10 c30 |
| Elf (archer) | 507 | r11 c23 (arrow at 531 = r12 c03, dissolve at 539 = r12 c11) |

Monsters - contiguous from r00 c00, each 24-tile walk (frame*8+dir, 3 frames):

| block | start | our kind |
|---|---|---|
| ghost | 0 | K_GHOST |
| grunt | 24 | K_GRUNT |
| floating head | 48 | K_THIEF (guess) |
| imp | 72 | K_DEMON (guess) |
| bad wizard | 96 | K_SORC (guess) |
| death | 120 | K_DEATH |
| blob | 144 | K_LOBBER (guess) |
| pulse (8) | 168 | spawn/other fx |
| blue grunt | 176 | (unused) |
| swordman | 200 | (unused) |

Effects: generator = **224** (r05 c04). demon/enemy fireball 8-dir = **323**
(r07 c15, +dir). Elf arrow 8-dir = **531** (r12 c03, +dir). r05 c05-c19 -
three unlabelled fx sets (225-230, 231-234, 235-239), tentatively lobber
rock / potion blast / death puff, confirm on playtest.
Items: meat = 282 (r06 c18), key = 283 (r06 c19), potion = 289 (r06 c25).

### Not yet mapped (needed for wall autotiling - M5c)

- **Thin blue wall autotile set** - `r12 c36-c43` + `r13 c00-c10` (~19 tiles;
  user says 15 orientation pieces + 2 diagonals + 1 blob). Orientation->tile
  mapping unknown.
- **Thick brown "packing slab" wall autotile set** - `r12 c16-c22`, `c27`,
  `c32`, `c33` (~10 visible; user says 15). Orientation->tile mapping unknown.

**Status:** M5a uses the solid wall block + floor + EXIT (confirmed).
Autotiling and real actor sprites/facing are M5b, pending the tile map.
