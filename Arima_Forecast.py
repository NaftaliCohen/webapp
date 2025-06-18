import os
import pyodbc
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.stattools import adfuller
from sklearn.metrics import mean_absolute_error, mean_squared_error
from collections import deque
from statistics import NormalDist

# הגדרות Tcl/Tk ל-Windows
os.environ['TCL_LIBRARY'] = r'C:\\Program Files\\Python313\\tcl\\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\\Program Files\\Python313\\tcl\\tk8.6'

# --- פונקציית סימולציית מלאי ---
def simulate_inventory(forecast_df, actual_df, lt_days, minL, maxL, initial_inventory, service_level=0.98):
    lt_weeks = max(1, int(np.ceil(lt_days / 7)))
    z = NormalDist().inv_cdf(service_level)

    merged = forecast_df.merge(actual_df, on="Date", how="left").fillna(0)
    merged["Error"] = merged["Actual"] - merged["Forecast"]
    err_sigma = merged["Error"].std(ddof=0)
    safety_stock = z * err_sigma * np.sqrt(lt_weeks)

    inventory = initial_inventory
    arrivals = deque([0.0] * lt_weeks)
    results = []

    for i in range(len(forecast_df)):
        forecast = forecast_df.loc[i, "Forecast"]
        inventory += arrivals.popleft()
        arrivals.append(0.0)

        inv_begin = inventory
        inventory -= forecast

        worst_demand = forecast_df["Forecast"].iloc[i:i+lt_weeks].sum()
        on_order = sum(arrivals)
        worst_inventory = inventory - worst_demand + on_order

        if worst_inventory < (minL + safety_stock):
            order_qty = (minL + safety_stock - worst_inventory) + (maxL - minL)
        else:
            order_qty = 0.0

        arrivals[-1] = order_qty
        inv_end = inventory

        results.append({
            "Date": forecast_df.loc[i, "Date"],
            "Forecast": forecast,
            "Inventory_Begin": inv_begin,
            "OrderQty": order_qty,
            "ArrivalQty": arrivals[0],
            "Inventory_End": inv_end
        })

    return pd.DataFrame(results)

# --- פונקציה לניתוח מצב מלאי ---
def analyze_inventory_status(inventory_df, min_level, max_level):
    inventory_df = inventory_df.copy()
    inventory_df["Status"] = inventory_df["Inventory_End"].apply(
        lambda x: "Below Min" if x < min_level else (
                  "Above Max" if x > max_level else "Within Range"))

    total = len(inventory_df)
    below = (inventory_df["Status"] == "Below Min").sum()
    within = (inventory_df["Status"] == "Within Range").sum()
    above = (inventory_df["Status"] == "Above Max").sum()

    print(f"\n📦 ניתוח סטטוס מלאי:")
    print(f"🔻 מתחת למינימום: {below / total * 100:.2f}%")
    print(f"✅ בתוך הטווח: {within / total * 100:.2f}%")
    print(f"🔺 מעל למקסימום: {above / total * 100:.2f}%")

    return inventory_df

# --- קריאת קובץ ---
df = pd.read_csv('optical_db_update.csv', encoding='utf-8')

df['ConsumptionDate'] = pd.to_datetime(df['ConsumptionDate'], errors='coerce')
df['ConsumptionQty'] = pd.to_numeric(df['ConsumptionQty'], errors='coerce')
df.dropna(subset=['ConsumptionDate', 'ConsumptionQty'], inplace=True)
df['Week'] = df['ConsumptionDate'].dt.to_period('W')

# סינון מק"טים עם לפחות 2 שבועות
valid_weeks = df[df['ConsumptionQty'] > 0].groupby('ItemCode')['Week'].nunique()
valid_items = valid_weeks[valid_weeks >= 2].index
df = df[df['ItemCode'].isin(valid_items)]

summary_stats = []  # לאיסוף נתוני סיכום לכל מק"ט

# לולאה על כל מק"ט
item_codes = df['ItemCode'].unique()
for item in item_codes:
    print(f"\n========== ItemCode: {item} ==========")
    item_df = df[df['ItemCode'] == item].copy()
    weekly = item_df.groupby('Week')['ConsumptionQty'].sum().to_timestamp()

    if len(weekly) <= 2:
        print("⛔ לא מספיק נתונים שבועיים – מדלג.")
        continue

    train = weekly[weekly.index.year <= 2023]
    test = weekly[weekly.index.year == 2024]

    if test.empty:
        print("⚠️ אין נתונים לשנת 2024. מדלג.")
        continue

    try:
        adf_result = adfuller(train)
        p_value = adf_result[1]
        d_value = 0 if p_value < 0.05 else 1
        print(f"p-value: {p_value:.4f} → d = {d_value}")
    except Exception as e:
        print(f"שגיאה בבדיקת תחנתיות: {e}")
        d_value = 1

    try:
        model = ARIMA(train, order=(1, d_value, 1))
        model_fit = model.fit()

        forecast = model_fit.forecast(steps=len(test))
        forecast.index = test.index

        mae = mean_absolute_error(test, forecast)
        rmse = np.sqrt(mean_squared_error(test, forecast))
        print(f"MAE: {mae:.2f}")
        print(f"RMSE: {rmse:.2f}")

        all_weeks = sorted(set(train.index.to_period('W')) |
                           set(test.index.to_period('W')) |
                           set(forecast.index.to_period('W')))
        week_labels = [w.strftime('%Y-W%U') for w in all_weeks]
        train_vals = [train.get(w.start_time, np.nan) for w in all_weeks]
        test_vals = [test.get(w.start_time, np.nan) for w in all_weeks]
        forecast_vals = [forecast.get(w.start_time, np.nan) for w in all_weeks]

        plt.figure(figsize=(14, 5))
        plt.plot(week_labels, test_vals, label='Actual Consumption 2024', marker='o')
        plt.plot(week_labels, forecast_vals, label='Forecast 2024', linestyle='--', marker='x')
        plt.title(f'Arima Weekly Forecast vs Actual for ItemCode {item}')
        plt.xlabel('Week')
        plt.ylabel('Quantity')
        plt.xticks(rotation=45)
        plt.legend()
        plt.grid(True, linestyle='--', alpha=0.5)
        plt.tight_layout()
        plt.show()

        # ⚙️ סימולציית מלאי
        forecast_df = pd.DataFrame({
            "Date": forecast.index,
            "Forecast": forecast.values
        })
        actual_df = pd.DataFrame({
            "Date": test.index,
            "Actual": test.values
        })

        lt_days = 21
        minL = 150
        maxL = 400
        initial_inventory = 300

        sim_result = simulate_inventory(forecast_df, actual_df, lt_days, minL, maxL, initial_inventory)

        plt.figure(figsize=(12, 4))
        plt.plot(sim_result["Date"], sim_result["Inventory_End"], label="Inventory")
        plt.axhline(minL, color="green", linestyle="--", label="Min Level")
        plt.axhline(maxL, color="red", linestyle="--", label="Max Level")
        plt.title(f"Inventory Simulation for ItemCode {item}")
        plt.ylabel("Units")
        plt.xticks(rotation=45)
        plt.grid(True, linestyle='--', alpha=0.5)
        plt.legend()
        plt.tight_layout()
        plt.show()

        # ניתוח סטטוס מלאי
        _ = analyze_inventory_status(sim_result, minL, maxL)
        analyzed_df = analyze_inventory_status(sim_result, minL, maxL)

        # אחוז בתוך תחום
        percent_within = (analyzed_df["Status"] == "Within Range").mean() * 100
        summary_stats.append({
            "ItemCode": item,
            "Percent Within Range": round(percent_within, 2)
        })

    except Exception as e:
        print(f"שגיאה במודל עבור מק\"ט {item}: {e}")

# --- סיכום כולל ---
summary_df = pd.DataFrame(summary_stats)
print("\n📊 סיכום אחוזים שכל מק\"ט היה בטווח התקין:")
print(summary_df.to_string(index=False))
