"""Dumbbell plot: prompt-based vs. transformer-based automated scoring (QWK).

Condensed, paper-ready version of the original R figure: wide aspect ratio,
large fonts, no title/subtitle/footnote (covered by the figure caption).
Outputs PNG (600 dpi) and PDF (vector) to the figures/ directory.
"""

import matplotlib.pyplot as plt
from matplotlib.lines import Line2D

# (study, comparison label, prompt QWK, transformer QWK), top row first
DATA = [
    ("S193", "S193-c", 0.87, 0.94),
    ("S193", "S193-b", 0.81, 0.87),
    ("S193", "S193-a", 0.61, 0.74),
    ("S189", "S189-b", 0.88, 0.85),
    ("S189", "S189-a", 0.84, 0.81),
    ("S187", "S187-b", 0.76, 0.77),
    ("S187", "S187-a", 0.74, 0.79),
    ("S168", "S168",   0.67, 0.74),
    ("S188", "S188",   0.50, 0.55),
]

PROMPT_COLOR = "#2b6cb8"       # blue
TRANSFORMER_COLOR = "#d9533f"  # red
BAND_COLOR = "#dce9f5"

TICK_FS = 18
LABEL_FS = 20
VALUE_FS = 16
LEGEND_FS = 17
GROUP_FS = 18
NOTE_FS = 16

fig, ax = plt.subplots(figsize=(15, 5.4))

n = len(DATA)
ys = list(range(n - 1, -1, -1))  # first row at the top

# Alternating shaded bands per study
studies = []
for study, *_ in DATA:
    if not studies or studies[-1][0] != study:
        studies.append([study, 1])
    else:
        studies[-1][1] += 1

row = 0
for i, (study, size) in enumerate(studies):
    y_top = ys[row] + 0.5
    y_bot = ys[row + size - 1] - 0.5
    if i % 2 == 0:
        ax.axhspan(y_bot, y_top, color=BAND_COLOR, alpha=0.55, zorder=0)
    # Bold study label left of the tick labels (skip single-row groups,
    # whose tick label already names the study)
    if size > 1:
        ax.text(-0.082, (y_top + y_bot) / 2, study, transform=ax.get_yaxis_transform(),
                ha="right", va="center", fontsize=GROUP_FS, fontweight="bold",
                color="#444444")
    row += size

# Dumbbells with value labels placed outward (left of the lower value,
# right of the higher) so close pairs never collide
for (study, label, p, t), y in zip(DATA, ys):
    ax.plot([p, t], [y, y], color="#b0b0b0", lw=2.5, zorder=2)
    ax.scatter([p], [y], s=210, color=PROMPT_COLOR, zorder=3)
    ax.scatter([t], [y], s=210, color=TRANSFORMER_COLOR, zorder=3)
    lo, hi = sorted([(p, PROMPT_COLOR), (t, TRANSFORMER_COLOR)])
    ax.text(lo[0] - 0.013, y, f"{lo[0]:.2f}", ha="right", va="center",
            fontsize=VALUE_FS, color=lo[1], fontweight="bold", zorder=4)
    ax.text(hi[0] + 0.013, y, f"{hi[0]:.2f}", ha="left", va="center",
            fontsize=VALUE_FS, color=hi[1], fontweight="bold", zorder=4)

ax.set_yticks(ys)
ax.set_yticklabels([d[1] for d in DATA], fontsize=TICK_FS)
for ticklabel, (study, size_label, *_ ) in zip(ax.get_yticklabels(), DATA):
    if study == size_label:  # single-row study: bold like a group label
        ticklabel.set_fontweight("bold")
        ticklabel.set_color("#444444")

ax.set_xlim(0.4, 1.0)
ax.set_ylim(-0.6, n - 0.4)
ax.set_xticks([0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0])
ax.tick_params(axis="x", labelsize=TICK_FS)
ax.set_xlabel("Quadratic Weighted Kappa (QWK)", fontsize=LABEL_FS)

ax.xaxis.grid(True, color="#d5d5d5", lw=0.8, zorder=0)
ax.set_axisbelow(True)
for side in ("top", "right", "left"):
    ax.spines[side].set_visible(False)
ax.tick_params(axis="y", length=0)

# Reader context note in the empty upper-left region
ax.text(0.015, 0.94, "Higher QWK = better agreement\nwith human raters",
        transform=ax.transAxes, ha="left", va="top", fontsize=NOTE_FS,
        color="#444444", style="italic", zorder=4,
        bbox=dict(boxstyle="round,pad=0.45", facecolor="white",
                  edgecolor="#b0b0b0", alpha=0.95))

legend_handles = [
    Line2D([], [], marker="o", linestyle="none", markersize=13,
           color=PROMPT_COLOR, label="Prompt"),
    Line2D([], [], marker="o", linestyle="none", markersize=13,
           color=TRANSFORMER_COLOR, label="Transformer"),
]
ax.legend(handles=legend_handles, title="Method", loc="lower right",
          fontsize=LEGEND_FS, title_fontsize=LEGEND_FS, framealpha=0.95,
          borderpad=0.6)

fig.tight_layout()
fig.savefig("figures/prompt_vs_transformer_qwk.png", dpi=600, bbox_inches="tight")
fig.savefig("figures/prompt_vs_transformer_qwk.pdf", bbox_inches="tight")
print("wrote figures/prompt_vs_transformer_qwk.{png,pdf}")
