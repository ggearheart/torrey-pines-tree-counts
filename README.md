# Torrey Pine Tree Counts Over Time

A small data-shaping pipeline for Torrey pine (*Pinus torreyana*) abundance at
**Torrey Pines State Natural Reserve**, spanning roughly the last **3,000 years**.
It turns one hand-maintained spreadsheet into ready-to-use slide materials:
pre-cut view tables, charts, and a starter PowerPoint deck.

## The data is hybrid

| Era | Points | Source | Nature |
|---|---|---|---|
| Deep past (~1,000–3,000 yrs ago) | pollen estimates | Cole & Wahl 2000, *Quaternary Research* 53(3):341–351 (Los Peñasquitos Lagoon sediment cores) | **estimates**, not counts |
| Historical (1888–1962) | early naturalists | Lemmon, Belle Angier, Fleming, Ledig | estimates |
| Modern | 2006 (Santos & Franklin), 2018 (Tierra Data) | field surveys | **manual counts** |
| Recent | 2022, 2026 | counts + planting / post-fire estimates | hybrid |

Charts keep this honest: open/grey markers = estimates, filled pine-green =
manual counts, teal squares = count+estimate hybrids.

> **Caveat:** the deep-time points are derived from *pollen abundance*, not tree
> headcounts, and carry real uncertainty. Cole & Wahl find Torrey pine present
> throughout the record but less numerous before ~2,100 years ago.

## Year convention

In the source spreadsheet:

- **positive `year_CE`** = real calendar year CE (e.g. `1888`, `2026`)
- **negative `year_CE`** = *years before present* (e.g. `-3000` = 3,000 years ago)

`build.py` normalizes both onto a single calendar axis and computes
`years_before_present` from **2026** (today).

## Usage

```bash
python3 build.py
```

Reads the newest `data/TPSNR_treecount_*.xlsx`, then regenerates:

```
views/   view_full_3000yr.csv · view_modern_era.csv
         view_last_decade.csv · view_estimate_vs_actual.csv
charts/  chart_full_3000yr.png (log time axis) · chart_modern_era.png
         chart_estimate_vs_actual.png
TPSNR_treecount_slides.pptx   (6-slide starter deck)
```

To update: edit the spreadsheet (add rows, or save a new
`TPSNR_treecount_version2.xlsx` — the script auto-picks the newest one) and
re-run `build.py`.

## Requirements

Python 3 with `openpyxl`, `matplotlib`, and `python-pptx`.

```bash
pip install openpyxl matplotlib python-pptx
```

## Source

Cole, K.L. & Wahl, E. (2000). *A Late Holocene Paleoecological Record from
Torrey Pines State Reserve, California.* Quaternary Research 53(3): 341–351.
<https://doi.org/10.1006/qres.1999.2121>
