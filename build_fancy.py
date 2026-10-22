#!/usr/bin/env python3
"""
Fancy build path for the Torrey pine deck.

Same data as build.py, but a California State Parks-inspired theme:
coastal-sky gradients, Torrey Pines sandstone bluffs, and stylized wind-swept
Torrey pine silhouettes behind the charts and slides. All artwork is drawn
procedurally (no external image files), so this regenerates cleanly.

    python build_fancy.py

Outputs:
    charts_fancy/   themed PNG charts
    TPSNR_treecount_slides_fancy.pptx
"""
import os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
from matplotlib.patches import Ellipse, Polygon
from matplotlib.lines import Line2D

from build import find_source, read_rows, normalize  # reuse data layer

# ---------------------------------------------------------------- palette ----
SKY_TOP = "#BcD7E6"
SKY_HORZ = "#EAF0E4"
SEA = "#5E8CA6"
BLUFF = "#C79A5B"      # Torrey Pines sandstone
BLUFF_SH = "#A87C41"
PARK_GREEN = "#38613C"
PINE_DK = "#274A2C"
GOLD = "#C0901F"       # California gold
GOLD_LT = "#E0B64E"
BARK = "#5B4126"
INK = "#2A2A22"
PAPER = "#FBF7EE"
GRIDC = "#CBB78F"

HERE = os.path.dirname(os.path.abspath(__file__))
CH = os.path.join(HERE, "charts_fancy")
PPTX = os.path.join(HERE, "TPSNR_treecount_slides_fancy.pptx")

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 12,
    "svg.fonttype": "none",
})

# marker semantics preserved, themed colors
STYLE = {
    "count":    dict(label="Manual count",     color=PARK_GREEN, marker="o", fill=True),
    "hybrid":   dict(label="Estimate + count", color=GOLD,       marker="s", fill=True),
    "estimate": dict(label="Estimate",         color=BARK,       marker="o", fill=False),
}


# ------------------------------------------------------------- artwork -------
def _vgrad(c_top, c_bot, h=300):
    t = np.linspace(0, 1, h)[:, None]
    top = np.array(mcolors.to_rgb(c_top))
    bot = np.array(mcolors.to_rgb(c_bot))
    arr = (1 - t) * top + t * bot
    return np.repeat(arr[:, None, :], 2, axis=1)


def draw_pine(ax, x, base_y, h, lean=0.12, color=PINE_DK, alpha=1.0, z=2):
    """Stylized wind-swept Torrey pine in 0..1 axis coords."""
    w = h * 0.62
    tip_x = x + lean * h
    # trunk (slightly leaning, tapered)
    trunk = Polygon([[x - w * 0.045, base_y],
                     [x + w * 0.045, base_y],
                     [tip_x + w * 0.02, base_y + h * 0.62],
                     [tip_x - w * 0.02, base_y + h * 0.62]],
                    closed=True, color=BARK, alpha=alpha, zorder=z)
    ax.add_patch(trunk)
    # open, irregular, flat-ish wind-sculpted canopy: overlapping ellipses
    cx, cy = tip_x, base_y + h * 0.72
    blobs = [
        (cx, cy, w * 1.15, h * 0.34),
        (cx + w * 0.42, cy + h * 0.05, w * 0.72, h * 0.26),
        (cx - w * 0.40, cy - h * 0.02, w * 0.66, h * 0.24),
        (cx + w * 0.20, cy + h * 0.20, w * 0.60, h * 0.22),
        (cx + w * 0.70, cy - h * 0.04, w * 0.48, h * 0.18),  # wind-blown trailing edge
    ]
    for bx, by, bw, bh in blobs:
        ax.add_patch(Ellipse((bx, by), bw, bh, facecolor=color,
                             edgecolor="none", alpha=alpha, zorder=z))


def scene_axes(fig, pine_alpha=0.22, with_sea=True):
    """Paint a faint coastal-bluff scene across the whole figure; return it."""
    bg = fig.add_axes([0, 0, 1, 1], zorder=0)
    bg.set_xlim(0, 1); bg.set_ylim(0, 1); bg.axis("off")
    bg.imshow(_vgrad(SKY_TOP, SKY_HORZ), extent=[0, 1, 0, 1],
              aspect="auto", origin="upper", zorder=0)
    if with_sea:
        bg.add_patch(Polygon([[0, 0.14], [1, 0.14], [1, 0.20], [0, 0.185]],
                     closed=True, color=SEA, alpha=0.30, zorder=1))
    # sandstone bluff
    bg.add_patch(Polygon([[0, 0], [0, 0.17], [0.28, 0.15], [0.55, 0.19],
                          [0.78, 0.14], [1, 0.17], [1, 0]],
                 closed=True, color=BLUFF, alpha=0.9, zorder=1))
    bg.add_patch(Polygon([[0, 0], [0, 0.11], [0.5, 0.09], [1, 0.12], [1, 0]],
                 closed=True, color=BLUFF_SH, alpha=0.6, zorder=1))
    # a few pines on the bluff, kept to the edges so the plot center stays clean
    draw_pine(bg, 0.07, 0.15, 0.30, lean=0.10, alpha=pine_alpha, z=2)
    draw_pine(bg, 0.15, 0.14, 0.22, lean=0.16, alpha=pine_alpha * 0.8, z=2)
    draw_pine(bg, 0.93, 0.16, 0.34, lean=-0.10, alpha=pine_alpha, z=2)
    draw_pine(bg, 0.85, 0.13, 0.20, lean=-0.14, alpha=pine_alpha * 0.8, z=2)
    return bg


def _legend(ax, kinds, loc="upper left"):
    handles = []
    for kind in ["count", "hybrid", "estimate"]:
        if kind not in kinds:
            continue
        st = STYLE[kind]
        handles.append(Line2D([0], [0], marker=st["marker"], linestyle="none",
                              markersize=10,
                              markerfacecolor=(st["color"] if st["fill"] else "none"),
                              markeredgecolor=st["color"], markeredgewidth=1.8,
                              label=st["label"]))
    leg = ax.legend(handles=handles, frameon=True, fontsize=11, loc=loc)
    leg.get_frame().set_facecolor("white")
    leg.get_frame().set_alpha(0.82)
    leg.get_frame().set_edgecolor(GRIDC)


def fancy_chart(rows, title, filename, xmode="year", xfilter=None,
                logx=False, annotate=False, legend_loc="upper left"):
    sub = [r for r in rows if (xfilter is None or xfilter(r))]
    sub = sorted(sub, key=lambda r: r["calendar_year"])

    FLOOR = 0.8
    def xval(r):
        if xmode == "bp":
            return max(r["years_before_present"], FLOOR)
        return r["calendar_year"]

    fig = plt.figure(figsize=(11, 6.19), dpi=150)
    scene_axes(fig)
    ax = fig.add_axes([0.095, 0.135, 0.875, 0.76], zorder=3)
    ax.set_facecolor((1, 1, 1, 0.72))

    # connector
    chron = sorted(sub, key=lambda r: (r["years_before_present"] if xmode == "bp"
                                       else -r["calendar_year"]), reverse=(xmode != "bp"))
    ax.plot([xval(r) for r in chron], [r["count"] for r in chron],
            "-" if not logx else ":", color=GOLD_LT, lw=2.0, zorder=4, alpha=0.9)

    for kind, st in STYLE.items():
        pts = [r for r in sub if r["kind"] == kind]
        if not pts:
            continue
        ax.scatter([xval(r) for r in pts], [r["count"] for r in pts], s=110, zorder=6,
                   marker=st["marker"],
                   facecolors=(st["color"] if st["fill"] else "white"),
                   edgecolors=st["color"], linewidths=2.0)

    if annotate:
        for r in sub:
            ax.annotate(f'{int(r["count"]):,}', (xval(r), r["count"]),
                        textcoords="offset points", xytext=(0, 12),
                        ha="center", fontsize=12, color=INK, weight="bold", zorder=7)

    if logx:
        ax.set_xscale("log")
        ax.set_xlim(4000, FLOOR * 0.8)
        ticks = [3000, 2000, 1000, 500, 200, 100, 50, 20, 10, FLOOR]
        ax.set_xticks(ticks)
        ax.set_xticklabels(["3000", "2000", "1000", "500", "200", "100",
                            "50", "20", "10", "now"])
        ax.minorticks_off()
        ax.set_xlabel("Years before present (2026)  —  log scale", color=INK)
    elif xmode == "year":
        xs = [r["calendar_year"] for r in sub]
        if len(xs) <= 7:
            ax.set_xticks(xs); ax.set_xticklabels([str(x) for x in xs])
            ax.set_xlim(min(xs) - 2, max(xs) + 2)
            ax.set_ylim(0, max(r["count"] for r in sub) * 1.18)
        ax.set_xlabel("Year (CE)", color=INK)

    ax.set_ylabel("Torrey pine count", color=INK)
    ax.set_title(title, fontsize=16, color=PINE_DK, weight="bold", pad=12)
    ax.grid(axis="y", color=GRIDC, lw=0.8, alpha=0.7)
    ax.set_axisbelow(True)
    ax.tick_params(colors=INK)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    for sp in ("left", "bottom"):
        ax.spines[sp].set_color(GRIDC)
    _legend(ax, {r["kind"] for r in sub}, loc=legend_loc)

    os.makedirs(CH, exist_ok=True)
    out = os.path.join(CH, filename)
    fig.savefig(out, facecolor=PAPER)
    plt.close(fig)
    return out


# ------------------------------------------------- slide background images ---
def make_backgrounds():
    os.makedirs(CH, exist_ok=True)
    # rich title scene
    fig = plt.figure(figsize=(13.333, 7.5), dpi=150)
    scene_axes(fig, pine_alpha=0.95, with_sea=True)
    p = os.path.join(CH, "bg_title.png")
    fig.savefig(p, facecolor=PAPER); plt.close(fig)
    # subtle content scene
    fig = plt.figure(figsize=(13.333, 7.5), dpi=150)
    scene_axes(fig, pine_alpha=0.16, with_sea=False)
    # soft paper wash over the top for legible content
    wash = fig.add_axes([0, 0, 1, 1], zorder=5); wash.axis("off")
    wash.add_patch(Polygon([[0, 0.17], [1, 0.17], [1, 1], [0, 1]], closed=True,
                           color=PAPER, alpha=0.55))
    p2 = os.path.join(CH, "bg_content.png")
    fig.savefig(p2, facecolor=PAPER); plt.close(fig)
    return p, p2


# ----------------------------------------------------------------- deck ------
def build_deck(rows, charts, bg_title, bg_content, src_name):
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

    def rgb(hexs):
        h = hexs.lstrip("#")
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

    GREEN, GOLDc, INKc, CREAM = rgb(PARK_GREEN), rgb(GOLD), rgb(INK), rgb(PAPER)

    prs = Presentation()
    prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    W, H = prs.slide_width, prs.slide_height
    blank = prs.slide_layouts[6]

    def bg(slide, img):
        slide.shapes.add_picture(img, 0, 0, width=W, height=H)

    def tb(slide, l, t, w, h):
        x = slide.shapes.add_textbox(l, t, w, h); x.text_frame.word_wrap = True
        return x.text_frame

    def run(p, text, size, color, bold=False, italic=False):
        r = p.add_run(); r.text = text
        f = r.font; f.size = Pt(size); f.bold = bold; f.italic = italic
        f.color.rgb = color; f.name = "Arial"

    def goldbar(slide, t):
        b = slide.shapes.add_shape(1, Inches(0.9), t, Inches(3.2), Inches(0.07))
        b.fill.solid(); b.fill.fore_color.rgb = GOLDc; b.line.fill.background()
        b.shadow.inherit = False

    # ---- title
    s = prs.slides.add_slide(blank); bg(s, bg_title)
    panel = s.shapes.add_shape(1, Inches(0.7), Inches(2.5), Inches(9.2), Inches(2.6))
    panel.fill.solid(); panel.fill.fore_color.rgb = CREAM
    panel.shadow.inherit = False
    panel.line.color.rgb = GOLDc; panel.line.width = Pt(1.5)
    tf = tb(s, Inches(1.0), Inches(2.75), Inches(8.6), Inches(2.2))
    run(tf.paragraphs[0], "Torrey Pine Abundance Over Time", 38, GREEN, bold=True)
    p2 = tf.add_paragraph(); run(p2, "Torrey Pines State Natural Reserve — a ~3,000-year record", 19, INKc)
    p3 = tf.add_paragraph(); run(p3, "Pollen-based estimates through modern surveys", 14, GOLDc, italic=True)

    # ---- chart slides
    meta = {
        "full": ("The full record",
                 "Present on the right (log time axis). Ancient points are pollen estimates with real "
                 "uncertainty; modern points are surveys."),
        "modern": ("Modern era (1880–2026)",
                   "A late-1800s low of a few hundred trees recovers to several thousand. Open markers "
                   "are estimates; filled are hard counts."),
        "est": ("Estimate vs. hard count",
                "Only 2006 and 2018 are full manual counts; recent totals blend counts with planting "
                "and post-fire estimates."),
        "p2000": ("2000 to present",
                  "The last ~25 years, labeled. A gentle decline from the 2000 estimate (5,000) to the "
                  "2026 post-fire figure (3,950)."),
    }
    for key, img in charts:
        title, cap = meta[key]
        s = prs.slides.add_slide(blank); bg(s, bg_content)
        tf = tb(s, Inches(0.6), Inches(0.3), Inches(12), Inches(0.8))
        run(tf.paragraphs[0], title, 26, GREEN, bold=True)
        goldbar(s, Inches(1.05))
        s.shapes.add_picture(img, Inches(1.5), Inches(1.35), width=Inches(10.3))
        cf = tb(s, Inches(0.9), Inches(6.7), Inches(11.5), Inches(0.7))
        run(cf.paragraphs[0], cap, 12, INKc, italic=True)

    # ---- sources
    s = prs.slides.add_slide(blank); bg(s, bg_content)
    tf = tb(s, Inches(0.6), Inches(0.35), Inches(12), Inches(0.8))
    run(tf.paragraphs[0], "Methods & sources", 26, GREEN, bold=True)
    goldbar(s, Inches(1.1))
    body = tb(s, Inches(0.9), Inches(1.5), Inches(11.4), Inches(5.2))
    bullets = [
        "Deep-time points (~1,000–3,000 yrs ago) are ESTIMATES from pollen abundance in Los "
        "Peñasquitos Lagoon sediment cores — not tree counts.",
        "Source: Cole, K.L. & Wahl, E. (2000), Quaternary Research 53(3):341–351. Torrey pine present "
        "throughout but less numerous before ~2,100 years ago.",
        "Historical points (1888–1962): early naturalists (Lemmon, Belle Angier, Fleming, Ledig).",
        "Modern: 2006 (Santos & Franklin) and 2018 (Tierra Data) are manual counts; 2022/2026 blend "
        "counts with planting / post-fire estimates.",
        "Years-before-present computed from 2026. Estimates shown as open markers to signal uncertainty.",
        f"Generated from {src_name} by build_fancy.py.",
    ]
    for i, b in enumerate(bullets):
        p = body.paragraphs[0] if i == 0 else body.add_paragraph()
        run(p, "•  " + b, 14, INKc); p.space_after = Pt(9)

    prs.save(PPTX)
    return PPTX


# ----------------------------------------------------------------- main ------
def main():
    src = find_source()
    rows = normalize(read_rows(src))
    bg_title, bg_content = make_backgrounds()
    charts = [
        ("full",  fancy_chart(rows, "Torrey pine abundance over the full record (~3,000 years)",
                              "fancy_full_3000yr.png", xmode="bp", logx=True)),
        ("modern", fancy_chart(rows, "Torrey pine counts, modern era (1880–2026)",
                              "fancy_modern_era.png", xmode="year",
                              xfilter=lambda r: r["calendar_year"] >= 1880)),
        ("est",   fancy_chart(rows, "Estimate vs. hard count — modern era",
                              "fancy_estimate_vs_actual.png", xmode="year",
                              xfilter=lambda r: r["calendar_year"] >= 1880, annotate=True)),
        ("p2000", fancy_chart(rows, "Torrey pine counts, 2000–present",
                              "fancy_2000_present.png", xmode="year",
                              xfilter=lambda r: r["calendar_year"] >= 2000,
                              annotate=True, legend_loc="lower left")),
    ]
    deck = build_deck(rows, charts, bg_title, bg_content, os.path.basename(src))
    print(f"Source : {os.path.basename(src)}  ({len(rows)} rows)")
    print("Charts : " + ", ".join(os.path.basename(c) for _, c in charts))
    print(f"Deck   : {os.path.basename(deck)}")


if __name__ == "__main__":
    main()
