import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

# Real column headers are on the 3rd row of the sheet
data = pd.read_excel("G:\\My Drive\\EU_ETS1\\verified_emissions_2025_en.xlsx", sheet_name="data", header=2)

# Czech Republic accounts (stationary installations + aircraft operators)
cz = data[data["REGISTRY_CODE"] == "CZ"]

# Total verified emissions per year ("Excluded" entries -> NaN, ignored in the sum)
em_cols = [c for c in cz.columns if c.startswith("VERIFIED_EMISSIONS_")]
cz_totals = (
    cz[em_cols].apply(pd.to_numeric, errors="coerce").sum()
    .rename(lambda c: int(c.rsplit("_", 1)[1]))
    .sort_index()
)
print((cz_totals / 1e6).round(2).rename("Mt CO2e"))

# Plot
fig, ax = plt.subplots(figsize=(9, 5))
ax.plot(cz_totals.index, cz_totals / 1e6, color="#2a78d6", linewidth=2, marker="o", markersize=5)

# Direct labels on the first and last points only
for yr, dx, ha in ((cz_totals.index[0], 8, "left"), (cz_totals.index[-1], 0, "center")):
    ax.annotate(f"{cz_totals[yr] / 1e6:.1f} Mt", (yr, cz_totals[yr] / 1e6),
                textcoords="offset points", xytext=(dx, 10 if dx == 0 else 2), ha=ha, fontsize=9, color="#52514e")
ax.set_ylim(0, cz_totals.max() / 1e6 * 1.1)

ax.set_title("Czech Republic – total verified emissions, EU ETS", loc="left", fontsize=13)
ax.set_ylabel("Mt CO₂e", color="#52514e")

ax.xaxis.set_major_locator(mticker.MultipleLocator(2))
ax.grid(axis="y", color="#e5e4e0", linewidth=0.8)
for s in ("top", "right", "left"):
    ax.spines[s].set_visible(False)
ax.tick_params(colors="#52514e", length=0)
fig.tight_layout()
fig.savefig("cz_emissions.png", dpi=150)
plt.show()
