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

Row contents (each row is an 8-direction walk cycle unless noted):

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
| 12 | goblins + thrown rocks, then **wall tiles** (brown / dark / orange / red / blue brick, X-marked exit blocks) |
| 13 | **floor + wall tiles** (blue brick variants: plain / cracked / edged, green tiled) |

**Status:** asset library ready. The renderer currently draws characters in
Excel cells; wiring these in needs colour-per-cell or image-Shape rendering —
milestone M5.
