import os
import pyodbc
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.stattools import adfuller
from sklearn.metrics import mean_absolute_error, mean_squared_error

# הגדרות Tcl/Tk אם צריך
os.environ['TCL_LIBRARY'] = r'C:\Program Files\Python313\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Program Files\Python313\tcl\tk8.6'

# --- התחברות למסד נתונים ---
server = '.'
database = 'MyFinalproVData'
driver = '{ODBC Driver 17 for SQL Server}'

conn_str = f"""
DRIVER={driver};
SERVER={server};
DATABASE={database};
Trusted_Connection=yes;
"""

conn = pyodbc.connect(conn_str)

# --- שאיבת נתונים ---
query = "SELECT ItemCode, DocDate, Quantity FROM dbo.consumptions"
df = pd.read_sql(query, conn)
conn.close()

# --- עיבוד ראשוני ---
df['DocDate'] = pd.to_datetime(df['DocDate'], errors='coerce')
df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')
df.dropna(subset=['DocDate', 'Quantity'], inplace=True)

# --- סינון לפי שבועות עם צריכה חיובית ---
df['Week'] = df['DocDate'].dt.to_period('W')
valid_weeks = df[df['Quantity'] >= 0].groupby('ItemCode')['Week'].nunique()
valid_items = valid_weeks[valid_weeks >= 0].index
df = df[df['ItemCode'].isin(valid_items)]

item_codes = df['ItemCode'].unique()

# --- תחזית לפי שבועות ---
for item in item_codes:
    print(f"\n========== ItemCode: {item} ==========")

    item_df = df[df['ItemCode'] == item].copy()
    weekly = item_df.groupby('Week')['Quantity'].sum().to_timestamp()

    if len(weekly) <=2:
        print("⛔ לא מספיק נתונים שבועיים – מדלג.")
        continue

    # פיצול נתונים
    train = weekly[weekly.index.year <= 2023]
    test = weekly[weekly.index.year == 2024]

    if len(test) == 0:
        print("⚠️ אין נתונים לשנת 2024. מדלג.")
        continue

    # בדיקת תחנתיות
    try:
        adf_result = adfuller(train)
        p_value = adf_result[1]
        d_value = 0 if p_value < 0.05 else 1
        print(f"p-value: {p_value:.4f} → d = {d_value}")
    except Exception as e:
        print(f"שגיאה בבדיקת תחנתיות: {e}")
        d_value = 1  # ברירת מחדל

    try:
        # בניית המודל
        model = ARIMA(train, order=(1, d_value, 1))
        model_fit = model.fit()

        forecast = model_fit.forecast(steps=len(test))
        forecast.index = test.index

        # חישוב שגיאות
        mae = mean_absolute_error(test, forecast)
        rmse = np.sqrt(mean_squared_error(test, forecast))
        print(f"MAE: {mae:.2f}")
        print(f"RMSE: {rmse:.2f}")

        # המרה לתוויות של שבוע בשנה (WW)
        train_labels = train.index.to_series().dt.strftime('%Y-W%U')
        test_labels = test.index.to_series().dt.strftime('%Y-W%U')
        forecast_labels = forecast.index.to_series().dt.strftime('%Y-W%U')

        plt.figure(figsize=(14, 5))
        plt.plot(train_labels, train.values, label='Train Data')
        plt.plot(test_labels, test.values, label='Actual 2024', marker='o')
        plt.plot(forecast_labels, forecast.values, label='Forecast 2024', linestyle='--', marker='x')
        plt.title(f'Weekly Forecast vs Actual for ItemCode {item}')
        plt.xlabel('Week Number')
        plt.ylabel('Quantity')
        plt.xticks(rotation=45)
        plt.legend()
        plt.grid(True)
        plt.tight_layout()
        plt.show()

    except Exception as e:
        print(f"שגיאה במודל עבור מק\"ט {item}: {e}")
