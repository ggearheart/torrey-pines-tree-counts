#!/usr/bin/env python3
"""
Torrey Pine tree-count pipeline.

Reads the newest data/TPSNR_treecount_*.xlsx, normalizes the year convention,
and regenerates:
  views/   pre-cut CSV tables for slide-building
  charts/  PNG charts (light theme, 16:9)
  TPSNR_treecount_slides.pptx   starter deck

Re-run after editing the spreadsheet:   python build.py

Year convention (as agreed):
  * positive year_CE  -> real calendar year CE (e.g. 1888, 2026)
  * negative year_CE  -> YEARS BEFORE PRESENT (e.g. -3000 = 3000 years ago)
  * present = 2026    -> years_before_present and calendar years derived from it
"""
import csv
import glob
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from openpyxl import load_workbook

# ----------------------------------------------------------------------------
PRESENT = 2026
HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "data")
VIEWS = os.path.join(HERE, "views")
CHARTS = os.path.join(HERE, "charts")
PPTX_OUT = os.path.join(HERE, "TPSNR_treecount_slides.pptx")

PINE = "#2E5A34"     # manual counts
TEAL = "#2A7F9E"     # estimate + count (hybrid)
MUTED = "#9E9E9E"    # pure estimates
INK = "#222222"
GRID = "#DDDDDD"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 12,
    "axes.edgecolor": "#888888",
    "axes.linewidth": 0.8,
    "svg.fonttype": "none",
})


# ----------------------------------------------------------------------------
def find_source():
    cands = sorted(glob.glob(os.path.join(DATA_DIR, "TPSNR_treecount_*.xlsx")))
    cands = [c for c in cands if not os.path.basename(c).startswith("~$")]
    if not cands:
        # fall back to any master.xlsx
        cands = glob.glob(os.path.join(DATA_DIR, "*.xlsx"))
    if not cands:
        sys.exit("No source .xlsx found in data/")
    # newest by mtime
    return max(cands, key=os.path.getmtime)


def norm(s):
    return str(s).strip().lower().replace(" ", "_") if s is not None else ""


def read_rows(path):
    wb = load_workbook(path, data_only=True)
    ws = wb["tree_counts"] if "tree_counts" in wb.sheetnames else wb[wb.sheetnames[0]]
    # locate header row (the one containing 'count')
    header_row = None
    for r in range(1, min(ws.max_row, 20) + 1):
        vals = [norm(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
        if "count" in vals:
            header_row = r
            headers = vals
            break
    if header_row is None:
        sys.exit("Could not find a header row containing 'count'.")

    col = {h: i + 1 for i, h in enumerate(headers) if h}
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        def g(name):
            return ws.cell(row=r, column=col[name]).value if name in col else None
        count = g("count")
        year = g("year_ce")
        if count is None and year is None:
            continue
        if count is None:
            continue
        rows.append({
            "year_raw": year,
            "count": float(count),
            "method_category": g("method_category") or g("method") or "",
            "location": g("location") or "",
            "notes": g("notes") or "",
        })
    return rows


def normalize(rows):
    out = []
    for d in rows:
        raw = d["year_raw"]
        if raw is None:
            continue
        raw = float(raw)
        if raw < 0:                         # years-before-present entry
            years_ago = -raw
            cal = PRESENT - years_ago
        else:                               # calendar-year entry
            cal = raw
            years_ago = PRESENT - raw
        d["calendar_year"] = int(round(cal))
        d["years_before_present"] = int(round(years_ago))
        d["year_label"] = year_label(cal)
        d["kind"] = classify(d["method_category"])
        out.append(d)
    out.sort(key=lambda x: x["calendar_year"])
    return out


def year_label(cal):
    cal = int(round(cal))
    if cal <= 0:
        return f"~{abs(cal)} BCE"
    return f"{cal} CE"


def classify(method):
    m = norm(method)
    if "manual" in m or m in ("count", "census", "modern_census", "decadal_survey"):
        return "count"
    if "plus" in m or ("estimate" in m and "count" in m):
        return "hybrid"
    return "estimate"


STYLE = {
    "count":    dict(label="Manual count",        color=PINE,  marker="o", fill=True),
    "hybrid":   dict(label="Estimate + count",    color=TEAL,  marker="s", fill=True),
    "estimate": dict(label="Estimate",            color=MUTED, marker="o", fill=False),
}


# ----------------------------------------------------------------------------
def write_views(rows):
    os.makedirs(VIEWS, exist_ok=True)
    cols = ["calendar_year", "year_label", "years_before_present", "count",
            "kind", "method_category", "location", "notes"]

    def dump(name, data):
        with open(os.path.join(VIEWS, name), "w", newline="") as f:
            w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
            w.writeheader()
            for d in data:
                w.writerow(d)

    dump("view_full_3000yr.csv", rows)
    dump("view_modern_era.csv", [r for r in rows if r["calendar_year"] >= 1880])
    dump("view_last_decade.csv", [r for r in rows if r["calendar_year"] >= PRESENT - 12])
    # estimate vs actual, flagged + sorted by kind then year
    order = {"count": 0, "hybrid": 1, "estimate": 2}
    dump("view_estimate_vs_actual.csv",
         sorted(rows, key=lambda r: (order.get(r["kind"], 9), r["calendar_year"])))
    return ["view_full_3000yr.csv", "view_modern_era.csv",
            "view_last_decade.csv", "view_estimate_vs_actual.csv"]


# ----------------------------------------------------------------------------
def _scatter(ax, rows, xkey):
    for kind, st in STYLE.items():
        pts = [r for r in rows if r["kind"] == kind]
        if not pts:
            continue
        x = [r[xkey] for r in pts]
        y = [r["count"] for r in pts]
        ax.scatter(x, y, s=90, zorder=3,
                   marker=st["marker"],
                   facecolors=(st["color"] if st["fill"] else "white"),
                   edgecolors=st["color"], linewidths=1.8)


def _legend(ax, kinds):
    handles = []
    for kind in ["count", "hybrid", "estimate"]:
        if kind not in kinds:
            continue
        st = STYLE[kind]
        handles.append(Line2D([0], [0], marker=st["marker"], linestyle="none",
                              markersize=10,
                              markerfacecolor=(st["color"] if st["fill"] else "white"),
                              markeredgecolor=st["color"], markeredgewidth=1.8,
                              label=st["label"]))
    ax.legend(handles=handles, frameon=False, fontsize=11, loc="upper left")


def chart_full(rows):
    # Log time axis so 3,000-yr-old points and the dense modern cluster are both legible.
    # Present (0 yrs ago) is floored to 0.8 so it sits on the log scale.
    FLOOR = 0.8

    def xpos(r):
        return max(r["years_before_present"], FLOOR)

    fig, ax = plt.subplots(figsize=(11, 6.19), dpi=150)
    chron = sorted(rows, key=lambda r: r["years_before_present"])
    ax.plot([xpos(r) for r in chron], [r["count"] for r in chron],
            ":", color="#BBBBBB", zorder=1, lw=1.2)
    for kind, st in STYLE.items():
        pts = [r for r in rows if r["kind"] == kind]
        if not pts:
            continue
        ax.scatter([xpos(r) for r in pts], [r["count"] for r in pts], s=90, zorder=3,
                   marker=st["marker"],
                   facecolors=(st["color"] if st["fill"] else "white"),
                   edgecolors=st["color"], linewidths=1.8)
    ax.set_xscale("log")
    ax.set_xlim(4000, FLOOR * 0.8)            # ancient on left, present on right
    ticks = [3000, 2000, 1000, 500, 200, 100, 50, 20, 10, FLOOR]
    ax.set_xticks(ticks)
    ax.set_xticklabels(["3000", "2000", "1000", "500", "200", "100", "50", "20", "10", "now"])
    ax.minorticks_off()
    ax.set_xlabel("Years before present (2026)  —  log scale")
    ax.set_ylabel("Torrey pine count")
    ax.set_title("Torrey pine abundance over the full record (~3,000 years)",
                 fontsize=15, color=INK, weight="bold", pad=12)
    ax.grid(axis="y", color=GRID, lw=0.8)
    ax.set_axisbelow(True)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    _legend(ax, {r["kind"] for r in rows})
    fig.text(0.5, 0.005,
             "Log time axis compresses the sparsely-sampled deep past. "
             "Dotted line is a visual aid, not measured data.",
             fontsize=9.5, color="#777777", ha="center")
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    out = os.path.join(CHARTS, "chart_full_3000yr.png")
    fig.savefig(out, facecolor="white")
    plt.close(fig)
    return out


def chart_modern(rows):
    sub = [r for r in rows if r["calendar_year"] >= 1880]
    fig, ax = plt.subplots(figsize=(11, 6.19), dpi=150)
    chron = sorted(sub, key=lambda r: r["calendar_year"])
    ax.plot([r["calendar_year"] for r in chron],
            [r["count"] for r in chron], "-", color="#CCCCCC", zorder=1, lw=1.5)
    _scatter(ax, sub, "calendar_year")
    ax.set_xlabel("Year (CE)")
    ax.set_ylabel("Torrey pine count")
    ax.set_title("Torrey pine counts, modern era (1880–2026)",
                 fontsize=15, color=INK, weight="bold", pad=12)
    ax.grid(axis="y", color=GRID, lw=0.8)
    ax.set_axisbelow(True)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    _legend(ax, {r["kind"] for r in sub})
    fig.tight_layout()
    out = os.path.join(CHARTS, "chart_modern_era.png")
    fig.savefig(out, facecolor="white")
    plt.close(fig)
    return out


def chart_est_vs_actual(rows):
    """Modern era, emphasising which points are hard counts vs estimates."""
    sub = [r for r in rows if r["calendar_year"] >= 1880]
    fig, ax = plt.subplots(figsize=(11, 6.19), dpi=150)
    _scatter(ax, sub, "calendar_year")
    for r in sub:
        ax.annotate(f'{int(r["count"]):,}',
                    (r["calendar_year"], r["count"]),
                    textcoords="offset points", xytext=(0, 9),
                    ha="center", fontsize=8.5, color="#555555")
    ax.set_xlabel("Year (CE)")
    ax.set_ylabel("Torrey pine count")
    ax.set_title("Estimate vs. hard count — modern era",
                 fontsize=15, color=INK, weight="bold", pad=12)
    ax.grid(axis="y", color=GRID, lw=0.8)
    ax.set_axisbelow(True)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    _legend(ax, {r["kind"] for r in sub})
    fig.tight_layout()
    out = os.path.join(CHARTS, "chart_estimate_vs_actual.png")
    fig.savefig(out, facecolor="white")
    plt.close(fig)
    return out


def build_charts(rows):
    os.makedirs(CHARTS, exist_ok=True)
    return [chart_full(rows), chart_modern(rows), chart_est_vs_actual(rows)]


# ----------------------------------------------------------------------------
def build_pptx(rows, charts, src_name):
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

    PINE_RGB = RGBColor(0x2E, 0x5A, 0x34)
    INK_RGB = RGBColor(0x22, 0x22, 0x22)
    GREY_RGB = RGBColor(0x77, 0x77, 0x77)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    W, H = prs.slide_width, prs.slide_height

    def textbox(slide, l, t, w, h):
        tb = slide.shapes.add_textbox(l, t, w, h)
        tb.text_frame.word_wrap = True
        return tb.text_frame

    def set_run(p, text, size, color, bold=False, italic=False):
        r = p.add_run(); r.text = text
        f = r.font; f.size = Pt(size); f.bold = bold; f.italic = italic
        f.color.rgb = color; f.name = "Arial"

    # ---- title slide
    s = prs.slides.add_slide(blank)
    bar = s.shapes.add_shape(1, 0, Inches(2.6), W, Inches(0.08))
    bar.fill.solid(); bar.fill.fore_color.rgb = PINE_RGB; bar.line.fill.background()
    tf = textbox(s, Inches(0.9), Inches(2.9), Inches(11.5), Inches(2.4))
    p = tf.paragraphs[0]; set_run(p, "Torrey Pine Abundance Over Time", 40, PINE_RGB, bold=True)
    p2 = tf.add_paragraph(); set_run(p2, "Torrey Pines State Natural Reserve — ~3,000-year record", 20, INK_RGB)
    p3 = tf.add_paragraph()
    set_run(p3, "Pollen-based estimates (Cole & Wahl 2000) through modern surveys", 14, GREY_RGB, italic=True)

    # ---- chart slides
    captions = {
        "chart_full_3000yr.png":
            "Present on the right. Ancient points are pollen-derived estimates with real uncertainty; "
            "modern points are surveys. Story: present throughout, far less numerous in the deep past.",
        "chart_modern_era.png":
            "The documented era. A late-1800s low of a few hundred trees recovers to several thousand "
            "by the 2000s. Open markers are estimates; filled are hard counts.",
        "chart_estimate_vs_actual.png":
            "Which numbers are measured vs. estimated. Only the 2006 and 2018 points are full manual counts; "
            "recent totals blend counts with planting/fire estimates.",
    }
    titles = {
        "chart_full_3000yr.png": "The full record",
        "chart_modern_era.png": "Modern era (1880–2026)",
        "chart_estimate_vs_actual.png": "Estimate vs. hard count",
    }
    for ch in charts:
        name = os.path.basename(ch)
        s = prs.slides.add_slide(blank)
        tf = textbox(s, Inches(0.6), Inches(0.35), Inches(12), Inches(0.8))
        set_run(tf.paragraphs[0], titles.get(name, name), 26, PINE_RGB, bold=True)
        s.shapes.add_picture(ch, Inches(1.4), Inches(1.3), width=Inches(10.5))
        cf = textbox(s, Inches(0.8), Inches(6.55), Inches(11.7), Inches(0.8))
        set_run(cf.paragraphs[0], captions.get(name, ""), 12, GREY_RGB, italic=True)

    # ---- data table slide
    s = prs.slides.add_slide(blank)
    tf = textbox(s, Inches(0.6), Inches(0.35), Inches(12), Inches(0.8))
    set_run(tf.paragraphs[0], "The data", 26, PINE_RGB, bold=True)
    hdr = ["Year", "Yrs ago", "Count", "Type", "Source"]
    nrows = len(rows) + 1
    tbl = s.shapes.add_table(nrows, len(hdr), Inches(1.0), Inches(1.25),
                             Inches(11.3), Inches(0.3) * nrows).table
    tbl.columns[0].width = Inches(1.6); tbl.columns[1].width = Inches(1.3)
    tbl.columns[2].width = Inches(1.4); tbl.columns[3].width = Inches(2.2)
    tbl.columns[4].width = Inches(4.8)
    for j, h in enumerate(hdr):
        c = tbl.cell(0, j); c.text = h
        for para in c.text_frame.paragraphs:
            for r in para.runs:
                r.font.bold = True; r.font.size = Pt(12); r.font.name = "Arial"
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        c.fill.solid(); c.fill.fore_color.rgb = PINE_RGB
    kind_name = {"count": "Manual count", "hybrid": "Estimate + count", "estimate": "Estimate"}
    for i, d in enumerate(sorted(rows, key=lambda r: r["calendar_year"], reverse=True), start=1):
        vals = [d["year_label"], f'{d["years_before_present"]:,}',
                f'{int(d["count"]):,}', kind_name.get(d["kind"], d["kind"]),
                str(d["notes"])[:60]]
        for j, v in enumerate(vals):
            c = tbl.cell(i, j); c.text = v
            for para in c.text_frame.paragraphs:
                for r in para.runs:
                    r.font.size = Pt(10.5); r.font.name = "Arial"; r.font.color.rgb = INK_RGB

    # ---- methods / sources slide
    s = prs.slides.add_slide(blank)
    tf = textbox(s, Inches(0.6), Inches(0.4), Inches(12), Inches(0.8))
    set_run(tf.paragraphs[0], "Methods & sources", 26, PINE_RGB, bold=True)
    body = textbox(s, Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.3))
    bullets = [
        "Deep-time points (≈2,000–3,000 yrs ago) are ESTIMATES derived from pollen "
        "abundance in Los Peñasquitos Lagoon sediment cores — not tree counts.",
        "Source: Cole, K.L. & Wahl, E. (2000), Quaternary Research 53(3):341–351. The record shows "
        "Torrey pine present throughout but less numerous before ~2,100 years ago.",
        "Historical points (1888–1962) are early naturalists' estimates (Lemmon, Belle Angier, "
        "Fleming, Ledig).",
        "Modern points: 2006 (Santos & Franklin) and 2018 (Tierra Data) are manual counts; "
        "2022 and 2026 blend counts with planting / post-fire estimates.",
        "Years-before-present computed from 2026. Estimates carry real uncertainty — shown as open "
        "markers, not precise values.",
        f"Generated from {src_name} by build.py — re-run after editing the spreadsheet.",
    ]
    for i, b in enumerate(bullets):
        p = body.paragraphs[0] if i == 0 else body.add_paragraph()
        set_run(p, "•  " + b, 14, INK_RGB)
        p.space_after = Pt(10)

    prs.save(PPTX_OUT)
    return PPTX_OUT


# ----------------------------------------------------------------------------
def main():
    src = find_source()
    rows = normalize(read_rows(src))
    if not rows:
        sys.exit("No data rows found.")
    views = write_views(rows)
    charts = build_charts(rows)
    pptx = build_pptx(rows, charts, os.path.basename(src))
    print(f"Source : {os.path.basename(src)}  ({len(rows)} rows)")
    print("Views  : " + ", ".join(views))
    print("Charts : " + ", ".join(os.path.basename(c) for c in charts))
    print(f"Deck   : {os.path.basename(pptx)}")


if __name__ == "__main__":
    main()
