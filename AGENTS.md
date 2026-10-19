# PlayerCard — project memory

## Layout
- `PlayerCard/Cards OVR/` — final cards, named `<OVR> - <Name> - <Type>.png`, 1345x784, hard-baked PNG (stats drawn into pixels).
- `PlayerCard/Cards/` — older/loose cards, `_Card data.csv` inside is STALE (out of sync with Data.xlsx).
- `PlayerCard/Data/Data.xlsx` — SOURCE OF TRUTH for all numbers. 87 sheets: `Trang tính1` (global formulas/weights), one sheet per player, `Thay đổi so với gốc` (change log).
- `PlayerCard/Avatar Gốc/`, `Claude outputs/` — reference/preview images.
- No generation script exists in repo. Cards were produced externally.

## Data.xlsx sheet structure (each player sheet, 39x26)
- Row 1: player name + stat group header (FIN, LSH, POS, HEA, CRO | SPA...).
- Rows 2-11: composite group -> sub-stat weights (cols B-F), and position -> OVR table (cols H=POS, I=OVR) for ST/CF/WF/AM/WM/CM/PV/WB/FB/CB.
- Rows 16-39: sub-stat ratings, col A=name col B=value. Row 39 = Total.
- Rows 17,19,21,23,25,27: composite rating rows. Col D=composite label (ATK/DIS/CRE/PHY/MEN/DEF), col K=computed value, col L=rounded value used on card.
- OVR on card = the position table value for the player's PRIMARY card position.
- JUM formula: `JUM_new = ROUND((JUM_gốc + height_cm - 100)/2, 0)`. Height/JUM_gốc editable inputs on `Trang tính1`.

## Card text positions (Cards OVR, 1345x784)
- Big OVR number: around x=359-556, y=169-296 (varies by digits).
- Stat label column: x≈130 (ATK/DIS/CRE/PHY/MEN/DEF), y≈388,438,487,536,585,634.
- Stat value column: x≈535 (right-aligned region), same y rows.
- Key-attribute labels: x≈622, y≈420/457/495.
- Key-attribute values: x≈748, y≈420/457.
- Position text (e.g. CM/PV): x≈854, y≈457.
- Card type banner (BEAST/PLAYMAKER/...): x≈110-290, y≈273.
- Player name: right side, x≈890-1290, y≈547/587.
- Type + FC at bottom-right: x≈955-1140, y≈644.

## How to update stats (when asked)
1. Read numbers from `Data/Data.xlsx` (source of truth), NOT from `_Card data.csv`.
2. Target ONLY `Cards OVR/*.png`. **Overwrite the old card in place; if the OVR changed, also rename the file `<old> - ...` -> `<new> - ...`** (filename carries the OVR). Delete the old-named file.
3. Edit pixels directly: cover old glyph area with background then draw new text matching existing font/color (auto-detected from pixels — no font/color spec known).
4. PNG is 24bpp RGB, opaque.

## Proven card-editing method (works — verified on Kroos 83->84)
Background art and every panel are **byte-identical between cards of the same playstyle**. Only the ink differs: big OVR digits, 6 stat numbers, 3 key-attribute numbers, player name, banner text. So per playstyle:
- Build a clean **background plate** for the stat column and the key column (mode colour per pixel across all cards of that playstyle).
- Harvest digit sprites (0-9) from donor cards of the same playstyle (patch 14x24 for stats, 10x20 for key; pick the sample with max ink).
- Clear the two columns with the plate, then stamp new digits right-aligned.
- Big OVR: copy the whole OVR region from a same-playstyle card that already shows the target number.

### Geometry (Cards OVR, 1345x784)
- Stat digit row tops: y = 394, 443, 492, 541, 590, 639 (step 49). Right edge of right-aligned group x=563, pitch 14, glyph 14x24.
- Key digit rows: y = 422, 459, 496 (step 37). Right edge x=767, pitch 10, glyph 10x20.
- Ink colour is cream `255,225,179` on dark bg; stat column bg ≈ `17,29,45`, key panel bg ≈ `32,28,42`.
- Key-attribute labels/order: top 3 sub-stats from the sheet (Data.xlsx), descending.
- The playstyle banner and position text on the card do NOT change.

### Status
- Only PlayMaker sprite sets built so far. Other playstyles need their own harvest pass (same script, filter by playstyle).
- Working scripts lived in `C:\Users\hungn\AppData\Local\Temp\opencode\` (masks.ps1, patches.ps1, render_kroos.ps1) — temp, may be gone; regenerate with the geometry above.


## Tooling available on this machine
- No Python (`python`/`py` not installed — only Windows Store stub). No ImageMagick, no dotnet, no LibreOffice.
- Excel COM available (`New-Object -ComObject Excel.Application`) for reading/writing xlsx.
- System.Drawing available in PowerShell for PNG read/write.
- Windows.Media.Ocr (WinRT) available via PowerShell for reading text off cards.
- Git repo at `D:\GiaiTri_DesignPlayerCard` (root), branch `main`.
