import pandas as pd
import matplotlib.pyplot as plt
from neuralprophet import NeuralProphet
import calendar

# === Step 1: Load and preprocess dataset ===
df_full = pd.read_excel(r"C:\Users\Hp\Desktop\Graduate Proj\Graduate Project FInal Files\Daily_Average_Demand_2020_2024.xlsx")
df_full["ds"] = pd.to_datetime(df_full["Date"])
df_full["y"] = df_full["Daily Average"]
df_full["Year"] = df_full["ds"].dt.year
df_full["Month"] = df_full["ds"].dt.month
df_full["Day"] = df_full["ds"].dt.day

# === Step 2: Setup for plotting and tracking forecasts ===
forecast_all_months = []
fig, axes = plt.subplots(nrows=4, ncols=3, figsize=(18, 14), sharex=False, sharey=True)
month_names = ['January', 'February', 'March', 'April', 'May', 'June',
               'July', 'August', 'September', 'October', 'November', 'December']
colors = {2020: "green", 2021: "red", 2022: "blue", 2023: "orange", 2024: "cyan"}

# === Step 3: Loop through each month, train + forecast ===
for month in range(1, 13):
    days_in_month = calendar.monthrange(2025, month)[1]

    # 1. Filter historical data for this month from 2020–2024
    df_month = df_full[(df_full["Month"] == month) & (df_full["Year"] < 2025)].copy()
    df_train = df_month[["ds", "y"]].copy()

    # 2. Train NeuralProphet on this month's data
    model = NeuralProphet(
        yearly_seasonality=True,
        weekly_seasonality=True,
        daily_seasonality=True,
        epochs=120,
        learning_rate=0.0015,
        batch_size=3
    )
    model.fit(df_train, freq='D')

    # 3. Generate a large future dataframe so it definitely covers up to 2025
    df_future = model.make_future_dataframe(df_train, periods=500, n_historic_predictions=False)
    forecast = model.predict(df_future)

    # 4. Extract and store forecast for just the current month of 2025
    forecast = forecast[(forecast["ds"].dt.year == 2025) & (forecast["ds"].dt.month == month)].copy()
    forecast["Month"] = forecast["ds"].dt.month
    forecast["Day"] = forecast["ds"].dt.day
    forecast["Year"] = 2025
    forecast_all_months.append(forecast[["ds", "yhat1", "Month", "Day", "Year"]])

    # === Step 4: Plotting ===
    ax = axes[(month - 1) // 3, (month - 1) % 3]

    # Plot historical years
    for year in range(2020, 2025):
        df_hist = df_full[(df_full["Month"] == month) & (df_full["Year"] == year)]
        ax.plot(df_hist["Day"], df_hist["y"], label=str(year), color=colors.get(year, 'gray'))

    # Plot forecast for 2025
    ax.plot(forecast["Day"], forecast["yhat1"], label="2025", color="black", linestyle='--', linewidth=2)

    ax.set_title(month_names[month - 1])
    ax.set_xlabel("Day")
    ax.set_ylabel("Load (MW)")
    ax.grid(True)
    ax.set_xticks(range(1, days_in_month + 1, 2))

# === Step 5: Final layout and show ===
plt.tight_layout()
plt.suptitle("📊 Monthly Load Forecasts for 2025 vs Historical (2020–2024)", fontsize=18, y=1.02)
plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.03), ncol=6)
plt.show()

# === Step 6: Combine forecasts if needed ===
forecast_2025_full = pd.concat(forecast_all_months, ignore_index=True)
print(forecast_2025_full.head())

# Optional: Save forecast
# forecast_2025_full.to_excel("Forecast_2025_Monthly_Daily.xlsx", index=False)
# print("\nForecast saved as 'Forecast_2025_Monthly_Daily.xlsx'.")