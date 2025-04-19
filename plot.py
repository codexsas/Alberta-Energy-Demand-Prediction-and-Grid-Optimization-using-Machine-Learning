import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import calendar

# Load data
df = pd.read_excel(r"C:\Users\Hp\Desktop\Graduate Proj\Graduate Project FInal Files\2020-2025_Demand_Full.xlsx")

# Parse date and extract parts
df["Date"] = pd.to_datetime(df["Date"])
df["Year"] = df["Date"].dt.year
df["Month"] = df["Date"].dt.month
df["Day"] = df["Date"].dt.day

# Configs
colors = {
    2020: "green", 2021: "red", 2022: "blue",
    2023: "orange", 2024: "cyan", 2025: "black"
}
month_names = list(calendar.month_name)[1:]

# ===== Plot 1: Actual Daily Demand =====
fig1, axes1 = plt.subplots(nrows=4, ncols=3, figsize=(18, 14), sharex=False, sharey=True)

for i, month in enumerate(range(1, 13)):
    ax = axes1[i // 3, i % 3]
    for year in range(2020, 2026):
        subset = df[(df["Month"] == month) & (df["Year"] == year)]
        if subset.empty:
            continue
        linestyle = '--' if year == 2025 else '-'
        ax.plot(subset["Day"], subset["Daily Average"], label=str(year),
                color=colors.get(year, 'gray'), linestyle=linestyle)
    ax.set_title(month_names[month - 1])
    ax.set_xlabel("Day")
    ax.set_ylabel("Load (MW)")
    ax.grid(True)
    ax.set_xticks(range(1, 32, 2))

fig1.suptitle("📊 Daily Average Load Trends (2020–2025)", fontsize=18, y=1.02)
fig1.tight_layout()
fig1.legend(loc="upper center", bbox_to_anchor=(0.5, -0.03), ncol=6)

# ===== Plot 2: Trend Lines Only =====
fig2, axes2 = plt.subplots(nrows=4, ncols=3, figsize=(18, 14), sharex=False, sharey=True)

for i, month in enumerate(range(1, 13)):
    ax = axes2[i // 3, i % 3]
    for year in range(2020, 2026):
        subset = df[(df["Month"] == month) & (df["Year"] == year)]
        if subset.empty:
            continue
        x = subset["Day"]
        y = subset["Daily Average"]
        if len(x) > 1:
            coeffs = np.polyfit(x, y, 1)
            trend = np.polyval(coeffs, x)
            ax.plot(x, trend, label=str(year), color=colors.get(year, 'gray'), linestyle='-', linewidth=2)
    ax.set_title(month_names[month - 1])
    ax.set_xlabel("Day")
    ax.set_ylabel("Trend Load (MW)")
    ax.grid(True)
    ax.set_xticks(range(1, 32, 2))

fig2.suptitle("📈 Daily Demand Trend Lines (2020–2025)", fontsize=18, y=1.02)
fig2.tight_layout()
fig2.legend(loc="upper center", bbox_to_anchor=(0.5, -0.03), ncol=6)

plt.show()
