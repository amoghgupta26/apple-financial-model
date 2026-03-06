"""
Chart generation for Apple Inc. Financial Model
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import os

os.makedirs("/home/claude/financial_model/charts", exist_ok=True)

# ── Style Config ──────────────────────────────────────────────────────────────
HIST_YEARS = [2019, 2020, 2021, 2022, 2023]
PROJ_YEARS = [2024, 2025, 2026, 2027, 2028]
ALL_YEARS  = HIST_YEARS + PROJ_YEARS

plt.rcParams.update({
    'font.family': 'DejaVu Sans',
    'axes.spines.top': False,
    'axes.spines.right': False,
    'axes.grid': True,
    'grid.alpha': 0.3,
    'grid.linestyle': '--',
    'figure.facecolor': 'white',
    'axes.facecolor': '#F8FBFF',
})

NAVY    = '#1F3864'
BLUE    = '#2E75B6'
LBLUE   = '#9DC3E6'
GREEN   = '#00B050'
LGREEN  = '#92D050'
YELLOW  = '#FFD700'
ORANGE  = '#FF8C00'
RED     = '#FF0000'
GREY    = '#7F7F7F'
PROJ_ALPHA = 0.65

def add_hist_proj_divider(ax, x_pos=4.5, ymin=0, ymax=1, label=True):
    ax.axvline(x=x_pos, color=GREY, linestyle=':', linewidth=1.2, alpha=0.7)
    if label:
        ylim = ax.get_ylim()
        mid_y = (ylim[0] + ylim[1]) * 0.95
        ax.text(x_pos - 0.1, mid_y, 'Historical', ha='right', va='top',
                fontsize=8, color=GREY, style='italic')
        ax.text(x_pos + 0.1, mid_y, 'Projected', ha='left', va='top',
                fontsize=8, color=BLUE, style='italic')

# ═══════════════════════════════════════════════════════════════════════════════
# Chart 1: Revenue & EBITDA Bridge
# ═══════════════════════════════════════════════════════════════════════════════
revenue_hist = [260174, 274515, 365817, 394328, 383285]
ebitda_hist  = [76477,  77344,  120233, 130541, 125820]

# Projected
rev_growth = [0.043, 0.070, 0.080, 0.075, 0.065]
rev_proj = [revenue_hist[-1]]
for g in rev_growth:
    rev_proj.append(rev_proj[-1] * (1 + g))
rev_proj = rev_proj[1:]

ebitda_margins_proj = [0.320, 0.323, 0.326, 0.328, 0.328]
ebitda_proj = [r * m for r, m in zip(rev_proj, ebitda_margins_proj)]

fig, ax = plt.subplots(figsize=(12, 6))
x = np.arange(len(ALL_YEARS))
width = 0.38

bars1 = ax.bar(x[:5] - width/2, [r/1000 for r in revenue_hist],
               width, color=NAVY, label='Revenue (Historical)', zorder=3)
bars2 = ax.bar(x[5:] - width/2, [r/1000 for r in rev_proj],
               width, color=NAVY, alpha=PROJ_ALPHA, label='Revenue (Projected)',
               linestyle='--', edgecolor=NAVY, zorder=3)

bars3 = ax.bar(x[:5] + width/2, [e/1000 for e in ebitda_hist],
               width, color=BLUE, label='EBITDA (Historical)', zorder=3)
bars4 = ax.bar(x[5:] + width/2, [e/1000 for e in ebitda_proj],
               width, color=BLUE, alpha=PROJ_ALPHA, label='EBITDA (Projected)',
               edgecolor=BLUE, zorder=3)

# Value labels on top
for bars in [bars1, bars2, bars3, bars4]:
    for bar in bars:
        h = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, h + 2,
                f'${h:.0f}B', ha='center', va='bottom', fontsize=7.5, color='#333')

ax.set_xticks(x)
ax.set_xticklabels([str(y) for y in ALL_YEARS[:5]] +
                   [f"{y}E" for y in ALL_YEARS[5:]])
ax.set_ylabel('$ Billions', fontsize=11)
ax.set_title("Apple Inc. — Revenue & EBITDA (FY2019–FY2028E)",
             fontsize=13, fontweight='bold', color=NAVY, pad=15)
ax.legend(loc='upper left', fontsize=9)
add_hist_proj_divider(ax)
ax.set_ylim(0, 650)
plt.tight_layout()
plt.savefig('/home/claude/financial_model/charts/01_revenue_ebitda.png', dpi=150, bbox_inches='tight')
plt.close()
print("Chart 1 saved")

# ═══════════════════════════════════════════════════════════════════════════════
# Chart 2: Margin Trends
# ═══════════════════════════════════════════════════════════════════════════════
gross_margins_h = [37.8, 38.2, 41.8, 43.3, 44.1]
ebitda_margins_h = [29.4, 28.2, 32.9, 33.1, 32.8]
net_margins_h    = [21.2, 20.9, 25.9, 25.3, 25.3]

gross_margins_p  = [44.3, 44.8, 45.2, 45.5, 45.5]
ebitda_margins_p = [32.0, 32.3, 32.6, 32.8, 32.8]
net_margins_p    = [26.2, 26.5, 26.8, 27.0, 27.0]

fig, ax = plt.subplots(figsize=(12, 5.5))
x = np.arange(len(ALL_YEARS))

ax.plot(x[:5], gross_margins_h,  'o-', color='#00B050', lw=2.5, ms=7, label='Gross Margin')
ax.plot(x[5:], gross_margins_p,  's--', color='#00B050', lw=2, ms=6, alpha=0.7)
ax.plot(x[:5], ebitda_margins_h, 'o-', color=NAVY, lw=2.5, ms=7, label='EBITDA Margin')
ax.plot(x[5:], ebitda_margins_p, 's--', color=NAVY, lw=2, ms=6, alpha=0.7)
ax.plot(x[:5], net_margins_h,    'o-', color=BLUE, lw=2.5, ms=7, label='Net Margin')
ax.plot(x[5:], net_margins_p,    's--', color=BLUE, lw=2, ms=6, alpha=0.7)

# Connect hist to proj
for hist_d, proj_d, color in [(gross_margins_h, gross_margins_p, '#00B050'),
                                (ebitda_margins_h, ebitda_margins_p, NAVY),
                                (net_margins_h, net_margins_p, BLUE)]:
    ax.plot([4, 5], [hist_d[-1], proj_d[0]], '--', color=color, lw=1.5, alpha=0.5)

ax.set_xticks(x)
ax.set_xticklabels([str(y) for y in HIST_YEARS] + [f"{y}E" for y in PROJ_YEARS])
ax.set_ylabel('Margin (%)', fontsize=11)
ax.set_title("Apple Inc. — Key Margin Trends (FY2019–FY2028E)",
             fontsize=13, fontweight='bold', color=NAVY, pad=15)
ax.legend(loc='lower right', fontsize=10)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.0f}%'))
add_hist_proj_divider(ax)
ax.set_ylim(15, 55)
plt.tight_layout()
plt.savefig('/home/claude/financial_model/charts/02_margin_trends.png', dpi=150, bbox_inches='tight')
plt.close()
print("Chart 2 saved")

# ═══════════════════════════════════════════════════════════════════════════════
# Chart 3: Free Cash Flow
# ═══════════════════════════════════════════════════════════════════════════════
fcf_hist = [58896, 73365, 92953, 111443, 99584]
fcf_proj = [105000, 114000, 124000, 133000, 140000]

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5.5))

# FCF absolute
x = np.arange(len(ALL_YEARS))
bars_h = ax1.bar(x[:5], [f/1000 for f in fcf_hist], color=BLUE, label='Historical', zorder=3)
bars_p = ax1.bar(x[5:], [f/1000 for f in fcf_proj], color=BLUE, alpha=PROJ_ALPHA,
                  edgecolor=BLUE, label='Projected', zorder=3)
for bars in [bars_h, bars_p]:
    for bar in bars:
        h = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2, h + 1, f'${h:.0f}B',
                 ha='center', va='bottom', fontsize=8.5)

ax1.set_xticks(x)
ax1.set_xticklabels([str(y) for y in HIST_YEARS] + [f"{y}E" for y in PROJ_YEARS])
ax1.set_ylabel('$ Billions', fontsize=11)
ax1.set_title('Free Cash Flow ($B)', fontsize=12, fontweight='bold', color=NAVY)
ax1.legend()
add_hist_proj_divider(ax1)
ax1.set_ylim(0, 175)

# FCF margin waterfall
fcf_margins_h = [f/r*100 for f,r in zip(fcf_hist, revenue_hist)]
fcf_margins_p = [f/r*100 for f,r in zip(fcf_proj, rev_proj)]
ax2.fill_between(range(5), fcf_margins_h, alpha=0.3, color=BLUE)
ax2.fill_between(range(5,10), fcf_margins_p, alpha=0.15, color=BLUE)
ax2.plot(range(5), fcf_margins_h, 'o-', color=NAVY, lw=2.5, ms=8)
ax2.plot(range(5, 10), fcf_margins_p, 's--', color=NAVY, lw=2, ms=7, alpha=0.8)
ax2.plot([4,5], [fcf_margins_h[-1], fcf_margins_p[0]], '--', color=GREY, lw=1)

for i, (m, yr) in enumerate(zip(fcf_margins_h, HIST_YEARS)):
    ax2.text(i, m + 0.5, f'{m:.1f}%', ha='center', fontsize=8.5)
for i, (m, yr) in enumerate(zip(fcf_margins_p, PROJ_YEARS)):
    ax2.text(5+i, m + 0.5, f'{m:.1f}%', ha='center', fontsize=8.5, color=BLUE)

ax2.set_xticks(range(10))
ax2.set_xticklabels([str(y) for y in HIST_YEARS] + [f"{y}E" for y in PROJ_YEARS])
ax2.set_ylabel('FCF Margin (%)', fontsize=11)
ax2.set_title('FCF Margin (%)', fontsize=12, fontweight='bold', color=NAVY)
ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.0f}%'))
ax2.axvline(x=4.5, color=GREY, linestyle=':', lw=1.2)
ax2.set_ylim(18, 42)

plt.suptitle("Apple Inc. — Free Cash Flow Analysis", fontsize=13, fontweight='bold',
             color=NAVY, y=1.02)
plt.tight_layout()
plt.savefig('/home/claude/financial_model/charts/03_free_cash_flow.png', dpi=150, bbox_inches='tight')
plt.close()
print("Chart 3 saved")

# ═══════════════════════════════════════════════════════════════════════════════
# Chart 4: DCF Waterfall — EV Bridge
# ═══════════════════════════════════════════════════════════════════════════════
components = ['PV of\nFCFs', 'PV of\nTerminal\nValue', 'Enterprise\nValue',
              'Less:\nNet Debt', 'Equity\nValue']
values = [450, 2280, 2730, -66, 2664]
colors_wf = [BLUE, NAVY, GREEN, RED, GREEN]

fig, ax = plt.subplots(figsize=(10, 6))
cumulative = 0
running = [0, 450, 0, 2730, 0]
heights = [450, 2280, 2730, 66, 2664]

bar_bottoms = [0, 450, 0, 2730 - 66, 0]
bar_heights = [450, 2280, 2730, 66, 2664]
bar_colors  = [BLUE, NAVY, GREEN, RED, GREEN]

for i, (bottom, height, color, label) in enumerate(
        zip(bar_bottoms, bar_heights, bar_colors, components)):
    bar = ax.bar(i, height, bottom=bottom if i != 4 else 0,
                 color=color, width=0.6, zorder=3, edgecolor='white', linewidth=1.5)
    top = bottom + height
    ax.text(i, top + 30, f'${height:,}B' if i != 3 else f'(${height}B)',
            ha='center', va='bottom', fontweight='bold', fontsize=10)

    # Connector lines
    if i in [0, 1, 3]:
        next_bottom = bar_bottoms[i+1] if i < 4 else 0
        ax.plot([i + 0.3, i + 0.7], [bottom + height, bottom + height],
                color=GREY, lw=1, linestyle='--', alpha=0.5)

ax.set_xticks(range(5))
ax.set_xticklabels(components, fontsize=10)
ax.set_ylabel('$ Billions', fontsize=11)
ax.set_title("Apple Inc. — DCF Enterprise Value Bridge\n(Base Case: WACC 10.0% | TGR 3.0%)",
             fontsize=13, fontweight='bold', color=NAVY, pad=15)
ax.set_ylim(0, 3200)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'${x:,.0f}B'))

# Implied per share annotation
ax.annotate('Implied Share Price:\n~$172/share',
            xy=(4, 2664), xytext=(3.2, 2900),
            arrowprops=dict(arrowstyle='->', color=NAVY, lw=2),
            fontsize=11, fontweight='bold', color=NAVY,
            bbox=dict(boxstyle='round,pad=0.4', facecolor=YELLOW, alpha=0.8))

plt.tight_layout()
plt.savefig('/home/claude/financial_model/charts/04_dcf_bridge.png', dpi=150, bbox_inches='tight')
plt.close()
print("Chart 4 saved")

# ═══════════════════════════════════════════════════════════════════════════════
# Chart 5: Sensitivity Heatmap
# ═══════════════════════════════════════════════════════════════════════════════
wacc_vals = [0.085, 0.090, 0.095, 0.100, 0.105, 0.110, 0.115]
tgr_vals  = [0.015, 0.020, 0.025, 0.030, 0.035, 0.040]

base_pv_fcf   = 450000
base_net_debt = -65716
shares        = 15441
terminal_ufcf = 113000

matrix = []
for tgr in tgr_vals:
    row = []
    for wacc in wacc_vals:
        if wacc <= tgr:
            row.append(np.nan)
        else:
            tv    = terminal_ufcf * (1 + tgr) / (wacc - tgr)
            pv_tv = tv / (1 + wacc) ** 5
            ev    = base_pv_fcf + pv_tv
            eq    = ev + base_net_debt
            price = eq / shares
            row.append(round(price, 1))
    matrix.append(row)

matrix = np.array(matrix)

fig, ax = plt.subplots(figsize=(12, 6))
from matplotlib.colors import LinearSegmentedColormap
cmap = LinearSegmentedColormap.from_list('traffic',
    [(0,'#FF4444'), (0.4,'#FFD700'), (0.65,'#92D050'), (1,'#00B050')])

masked = np.ma.masked_invalid(matrix)
im = ax.imshow(masked, cmap=cmap, aspect='auto', vmin=80, vmax=280)

for i in range(len(tgr_vals)):
    for j in range(len(wacc_vals)):
        val = matrix[i, j]
        if not np.isnan(val):
            color = 'white' if val > 220 or val < 120 else 'black'
            weight = 'bold' if (j == 3 and i == 3) else 'normal'  # base case
            border = '★ ' if (j == 3 and i == 3) else ''
            ax.text(j, i, f'{border}${val:.0f}', ha='center', va='center',
                    color=color, fontsize=9.5, fontweight=weight)

ax.set_xticks(range(len(wacc_vals)))
ax.set_xticklabels([f'{w*100:.1f}%' for w in wacc_vals], fontsize=10)
ax.set_yticks(range(len(tgr_vals)))
ax.set_yticklabels([f'{g*100:.1f}%' for g in tgr_vals], fontsize=10)
ax.set_xlabel('WACC', fontsize=12, fontweight='bold')
ax.set_ylabel('Terminal Growth Rate', fontsize=12, fontweight='bold')
ax.set_title("Apple Inc. — DCF Sensitivity Analysis\nImplied Share Price: WACC × Terminal Growth Rate  (★ = Base Case)",
             fontsize=13, fontweight='bold', color=NAVY, pad=15)

cbar = plt.colorbar(im, ax=ax, shrink=0.8)
cbar.set_label('Implied Share Price ($)', fontsize=10)

# Mark base case
ax.add_patch(plt.Rectangle((2.5, 2.5), 1, 1, fill=False,
                             edgecolor='black', lw=3))

plt.tight_layout()
plt.savefig('/home/claude/financial_model/charts/05_sensitivity_heatmap.png', dpi=150, bbox_inches='tight')
plt.close()
print("Chart 5 saved")

# ═══════════════════════════════════════════════════════════════════════════════
# Chart 6: Trading Comps Scatter + Valuation Range
# ═══════════════════════════════════════════════════════════════════════════════
companies   = ['Apple\n(AAPL)', 'Microsoft\n(MSFT)', 'Alphabet\n(GOOGL)',
               'Meta\n(META)', 'Samsung\n(005930)', 'Sony\n(SONY)']
ev_ebitda   = [23.2, 30.4, 21.6, 23.1, 8.4, 14.9]
net_margins = [25.1, 35.0, 24.0, 29.0, 7.3, 7.8]
mkt_caps    = [2850, 3100, 2050, 1350, 380, 115]
colors_c    = [YELLOW, NAVY, BLUE, LBLUE, GREEN, GREY]

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

# Bubble chart
sc = ax1.scatter(net_margins, ev_ebitda,
                 s=[m * 0.4 for m in mkt_caps],
                 c=colors_c, alpha=0.85, edgecolors='white', linewidths=2, zorder=3)

for i, (comp, x, y) in enumerate(zip(companies, net_margins, ev_ebitda)):
    offset = (1.5, 0.8) if i == 0 else (1, 0.5)
    ax1.annotate(comp, (x, y), xytext=(x + offset[0], y + offset[1]),
                fontsize=8.5, ha='left',
                arrowprops=dict(arrowstyle='-', color=GREY, lw=0.8) if i > 0 else None)

ax1.set_xlabel('Net Profit Margin (%)', fontsize=11)
ax1.set_ylabel('EV / EBITDA Multiple', fontsize=11)
ax1.set_title('Peer Benchmarking\nEV/EBITDA vs. Net Margin (bubble = market cap)',
              fontsize=11, fontweight='bold', color=NAVY)
ax1.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.0f}%'))

# Valuation range bar chart
val_methods = ['DCF\n(Gordon)', 'DCF\n(Exit Mult.)', 'EV/Revenue\nComps', 'EV/EBITDA\nComps', 'Current\nPrice']
low_vals  = [145, 155, 140, 155, 195]
high_vals = [200, 210, 180, 185, 195]
mid_vals  = [(l+h)//2 for l,h in zip(low_vals, high_vals)]

y_pos = np.arange(len(val_methods))
colors_r = [NAVY, BLUE, GREEN, LBLUE, ORANGE]

for i, (low, high, mid, color, label) in enumerate(
        zip(low_vals, high_vals, mid_vals, colors_r, val_methods)):
    ax2.barh(i, high - low, left=low, height=0.5,
             color=color, alpha=0.8, zorder=3)
    ax2.plot(mid, i, 'D', color='white', ms=8, zorder=4)
    ax2.text(high + 2, i, f'${low}–${high}', va='center', fontsize=9)

ax2.set_yticks(y_pos)
ax2.set_yticklabels(val_methods, fontsize=9.5)
ax2.set_xlabel('Share Price ($)', fontsize=11)
ax2.set_title('Valuation Range Summary\nby Method', fontsize=11, fontweight='bold', color=NAVY)
ax2.set_xlim(100, 240)
ax2.axvline(x=195, color=ORANGE, linestyle='--', lw=2, alpha=0.8, label='Current Price $195')
ax2.legend(fontsize=9, loc='lower right')

plt.suptitle("Apple Inc. — Trading Comps & Valuation Summary", fontsize=13,
             fontweight='bold', color=NAVY, y=1.01)
plt.tight_layout()
plt.savefig('/home/claude/financial_model/charts/06_comps_valuation.png', dpi=150, bbox_inches='tight')
plt.close()
print("Chart 6 saved")

# ═══════════════════════════════════════════════════════════════════════════════
# Chart 7: Revenue Breakdown Estimate (Product vs Services)
# ═══════════════════════════════════════════════════════════════════════════════
years_seg = ['FY2021', 'FY2022', 'FY2023', 'FY2024E', 'FY2025E']
products  = [297392, 316199, 298085, 308000, 322000]
services  = [68425,  78129,  85200,  93000,  103000]

fig, ax = plt.subplots(figsize=(11, 5.5))
x = np.arange(len(years_seg))
width = 0.55

b1 = ax.bar(x, [p/1000 for p in products], width, label='Products', color=NAVY, zorder=3)
b2 = ax.bar(x, [s/1000 for s in services], width,
            bottom=[p/1000 for p in products], label='Services', color=BLUE, zorder=3)

# Value labels
for i, (p, s) in enumerate(zip(products, services)):
    total = (p + s) / 1000
    ax.text(i, total + 2, f'${total:.0f}B', ha='center', fontweight='bold', fontsize=9.5)
    ax.text(i, p/1000/2, f'{p/1000:.0f}B', ha='center', va='center',
            color='white', fontsize=8.5)
    ax.text(i, p/1000 + s/1000/2, f'{s/1000:.0f}B', ha='center', va='center',
            color='white', fontsize=8.5)

# Services growth annotation
svc_pct = [s/(p+s)*100 for p, s in zip(products, services)]
ax2 = ax.twinx()
ax2.plot(x, svc_pct, 'o--', color=GREEN, lw=2, ms=8, label='Services %')
ax2.set_ylabel('Services % of Revenue', color=GREEN, fontsize=10)
ax2.tick_params(axis='y', labelcolor=GREEN)
ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.0f}%'))
ax2.set_ylim(0, 40)

ax.set_xticks(x)
ax.set_xticklabels(years_seg)
ax.set_ylabel('$ Billions', fontsize=11)
ax.set_title("Apple Inc. — Revenue Segments: Products vs. Services",
             fontsize=13, fontweight='bold', color=NAVY, pad=15)

lines1, labels1 = ax.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax.legend(lines1 + lines2, labels1 + labels2, loc='upper left', fontsize=10)
ax.set_ylim(0, 500)

plt.tight_layout()
plt.savefig('/home/claude/financial_model/charts/07_revenue_segments.png', dpi=150, bbox_inches='tight')
plt.close()
print("Chart 7 saved")

# ═══════════════════════════════════════════════════════════════════════════════
# Chart 8: Balance Sheet Composition
# ═══════════════════════════════════════════════════════════════════════════════
bs_years = ['FY2019', 'FY2020', 'FY2021', 'FY2022', 'FY2023']
cash     = [100557, 90943, 90338, 48304, 61555]
ppe      = [37378,  36766, 39440, 42117, 43715]
other_a  = [200581, 196179, 221224, 262334, 247313]
lt_debt  = [91807,  98667, 109106, 98959, 95281]
equity   = [90488,  65339, 63090, 50672, 62146]
other_l  = [156221, 159882, 178716, 203124, 195156]

fig, axes = plt.subplots(1, 2, figsize=(14, 6))

x = np.arange(len(bs_years))
w = 0.55

# Assets stacked
axes[0].bar(x, [c/1000 for c in cash], w, label='Cash & Investments', color='#00B050', zorder=3)
axes[0].bar(x, [p/1000 for p in ppe], w, bottom=[c/1000 for c in cash],
            label='PP&E', color=BLUE, zorder=3)
axes[0].bar(x, [o/1000 for o in other_a], w,
            bottom=[(c+p)/1000 for c,p in zip(cash, ppe)],
            label='Other Assets', color=LBLUE, zorder=3)
axes[0].set_title('Total Assets Composition ($B)', fontsize=11, fontweight='bold', color=NAVY)
axes[0].set_xticks(x); axes[0].set_xticklabels(bs_years, rotation=30)
axes[0].set_ylabel('$ Billions'); axes[0].legend(fontsize=8.5)

# Liabilities + Equity stacked
axes[1].bar(x, [l/1000 for l in lt_debt], w, label='Long-Term Debt', color=RED, zorder=3)
axes[1].bar(x, [o/1000 for o in other_l], w,
            bottom=[l/1000 for l in lt_debt], label='Other Liabilities', color=ORANGE, zorder=3)
axes[1].bar(x, [e/1000 for e in equity], w,
            bottom=[(l+o)/1000 for l,o in zip(lt_debt, other_l)],
            label='Total Equity', color=GREEN, zorder=3)
axes[1].set_title('Liabilities & Equity Composition ($B)', fontsize=11, fontweight='bold', color=NAVY)
axes[1].set_xticks(x); axes[1].set_xticklabels(bs_years, rotation=30)
axes[1].set_ylabel('$ Billions'); axes[1].legend(fontsize=8.5)

plt.suptitle("Apple Inc. — Balance Sheet Evolution (FY2019–FY2023)",
             fontsize=13, fontweight='bold', color=NAVY, y=1.01)
plt.tight_layout()
plt.savefig('/home/claude/financial_model/charts/08_balance_sheet.png', dpi=150, bbox_inches='tight')
plt.close()
print("Chart 8 saved")

print("\n✅ All 8 charts generated successfully!")
