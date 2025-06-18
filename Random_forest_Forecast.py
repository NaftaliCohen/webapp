import os
import pandas as pd
import numpy as np
from sklearn.ensemble import RandomForestRegressor
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_absolute_error, r2_score, mean_squared_error
from sklearn.preprocessing import LabelEncoder
from datetime import datetime, timedelta
import matplotlib.pyplot as plt

# הגדרות Tcl/Tk ל-Windows
os.environ['TCL_LIBRARY'] = r'C:\\Program Files\\Python313\\tcl\\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\\Program Files\\Python313\\tcl\\tk8.6'


# --- הגדרת פונקציות עזר לסימולציית מלאי ---
def simulate_inventory(forecast_df, actual_df, lt_days, minL, maxL, initial_inventory):
    """
    מבצעת סימולציית מלאי שבועית.

    Parameters:
    - forecast_df (pd.DataFrame): DataFrame עם עמודות 'Date' ו-'Forecast'.
    - actual_df (pd.DataFrame): DataFrame עם עמודות 'Date' ו-'Actual'.
    - lt_days (int): Lead Time בימים.
    - minL (int): רמת מלאי מינימלית.
    - maxL (int): רמת מלאי מקסימלית.
    - initial_inventory (int): מלאי התחלתי.

    Returns:
    - pd.DataFrame: DataFrame עם תוצאות הסימולציה.
    """
    sim_data = []
    current_inventory = initial_inventory
    orders = {}  # {order_arrival_date: quantity}

    # ודא שהתאריכים הם אובייקטי datetime
    forecast_df['Date'] = pd.to_datetime(forecast_df['Date'])
    actual_df['Date'] = pd.to_datetime(actual_df['Date'])

    # מיזוג תחזית וצריכה בפועל לפי תאריך
    # נניח ש- 'Date' הוא תאריך ההתחלה של השבוע
    merged_df = pd.merge(forecast_df, actual_df, on='Date', how='outer').sort_values('Date').reset_index(drop=True)
    merged_df['Actual'] = merged_df['Actual'].fillna(0)  # למקרה של תאריכים ללא צריכה בפועל

    # המרת lead time משבועות לימים לצורך חישוב מדויק יותר (אם הנתונים שבועיים)
    lt_weeks = int(np.ceil(lt_days / 7)) # המר במפורש ל-int של פייתון

    for index, row in merged_df.iterrows():
        current_date = row['Date']
        forecast_qty = row['Forecast']
        actual_qty = row['Actual']  # הצריכה בפועל לשבוע זה

        # 1. קליטת הזמנות שהגיעו השבוע
        orders_arrived_this_week = 0
        keys_to_remove = []
        for order_date, qty in orders.items():
            if order_date <= current_date:
                orders_arrived_this_week += qty
                keys_to_remove.append(order_date)
        for key in keys_to_remove:
            del orders[key]

        current_inventory += orders_arrived_this_week

        # 2. צריכה בפועל
        current_inventory -= actual_qty

        # 3. קבלת החלטת הזמנה (בהתבסס על תחזית)
        order_qty = 0
        if current_inventory < minL:
            # נסה להזמין כדי להגיע ל-maxL, תוך התחשבות בתחזית
            # אסטרטגיית הזמנה פשוטה: הזמן עד לרמת המקסימום
            order_qty = maxL - current_inventory
            if order_qty < 0:  # לוודא שלא מזמינים שלילי
                order_qty = 0

            if order_qty > 0:
                # תאריך הגעת ההזמנה הוא 'current_date' בתוספת LT (בשבועות)
                order_arrival_date = current_date + timedelta(weeks=lt_weeks)
                orders[order_arrival_date] = orders.get(order_arrival_date, 0) + order_qty

        sim_data.append({
            "Date": current_date,
            "Inventory_Start": current_inventory + actual_qty - orders_arrived_this_week,
            # מלאי לפני צריכה וקבלת הזמנות
            "Actual_Consumption": actual_qty,
            "Forecast_Consumption": forecast_qty,
            "Orders_Placed": order_qty,
            "Orders_Arrived": orders_arrived_this_week,
            "Inventory_End": current_inventory
        })

    return pd.DataFrame(sim_data)


def analyze_inventory_status(sim_result_df, minL, maxL):
    """
    מנתחת את סטטוס המלאי (מתחת למינימום, בטווח, מעל המקסימום).

    Parameters:
    - sim_result_df (pd.DataFrame): DataFrame עם תוצאות הסימולציה, כולל 'Inventory_End'.
    - minL (int): רמת מלאי מינימלית.
    - maxL (int): רמת מלאי מקסימלית.

    Returns:
    - pd.DataFrame: DataFrame עם עמודת 'Status' חדשה.
    """
    conditions = [
        sim_result_df["Inventory_End"] < minL,
        (sim_result_df["Inventory_End"] >= minL) & (sim_result_df["Inventory_End"] <= maxL),
        sim_result_df["Inventory_End"] > maxL
    ]
    choices = ["Below Min", "Within Range", "Above Max"]
    sim_result_df["Status"] = np.select(conditions, choices, default="Unknown")
    return sim_result_df


# --- קריאת הנתונים ועיבודם (כפי שהיה בקוד שלך) ---
df = pd.read_csv("optical_db_update.csv", encoding='utf-8')
df['ConsumptionDate'] = pd.to_datetime(df['ConsumptionDate'], errors='coerce')
df = df.dropna(subset=['ConsumptionQty', 'ConsumptionDate'])

# יצירת פיצ'רים מהתאריכים
df['ConsumptionYear'] = df['ConsumptionDate'].dt.year
df['ConsumptionWeek'] = df['ConsumptionDate'].dt.isocalendar().week
df['ConsumptionWeekday'] = df['ConsumptionDate'].dt.weekday
df['YearWeek'] = df['ConsumptionDate'].dt.strftime('%Y-%U')

# קידוד מק"ט
le = LabelEncoder()
df['ItemCode_encoded'] = le.fit_transform(df['ItemCode'].astype(str))

# קיבוץ לפי מק"ט ושבוע
grouped = df.groupby(['ItemCode', 'YearWeek']).agg({
    'ConsumptionQty': 'sum',
    'PRICE': 'mean',
    'LT_Days': 'mean',
    'InventoryBalance': 'mean',
    'ConsumptionYear': 'first',
    'ConsumptionWeek': 'first',
    'ItemCode_encoded': 'first'
}).reset_index()

# מילוי חסרים
grouped.fillna(method='ffill', inplace=True)

# הגדרת מודל
features = ['PRICE', 'LT_Days', 'InventoryBalance', 'ConsumptionYear', 'ConsumptionWeek', 'ItemCode_encoded']
X = grouped[features]
y = grouped['ConsumptionQty']

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
model = RandomForestRegressor(n_estimators=100, random_state=42)
model.fit(X_train, y_train)

# -----------------------------
# יצירת תחזית ל-52 שבועות קדימה (כפי שהיה בקוד שלך)
# -----------------------------
# ... (הקוד הזה נשאר זהה לקוד המקורי שלך, מטעמי קיצור לא חוזר עליו) ...
# מק"טים ייחודיים
item_codes = grouped['ItemCode'].unique()

# תאריך שבוע אחרון
last_date = df['ConsumptionDate'].max()
future_data = []

for item in item_codes:
    item_data = grouped[grouped['ItemCode'] == item].sort_values('YearWeek').iloc[-1]

    price = item_data['PRICE']
    lt_days_val = item_data['LT_Days']  # שינוי שם המשתנה כדי לא להתנגש עם lt_days של הסימולציה
    inventory = item_data['InventoryBalance']
    item_encoded = item_data['ItemCode_encoded']

    for i in range(1, 53):  # 53 שבועות
        future_week_date = last_date + timedelta(weeks=i)
        year = future_week_date.year
        week = future_week_date.isocalendar().week
        year_week = f"{year}-{str(week).zfill(2)}"

        future_data.append({
            'ItemCode': item,
            'YearWeek': year_week,
            'PRICE': price,
            'LT_Days': lt_days_val,
            'InventoryBalance': inventory,
            'ConsumptionYear': year,
            'ConsumptionWeek': week,
            'ItemCode_encoded': item_encoded
        })

# המרה ל-DataFrame
future_df = pd.DataFrame(future_data)

# חיזוי
X_future = future_df[features]
future_df['PredictedQty'] = model.predict(X_future)

# תוצאה סופית
future_forecast = future_df[['ItemCode', 'YearWeek', 'PredictedQty']]
print("--- תחזית ל-52 שבועות קדימה (דוגמה) ---")
print(future_forecast.head(20))

# --- סינון נתונים לשנת 2024 והערכת מודל (כפי שהיה בקוד שלך) ---
actual_2024 = grouped[grouped['ConsumptionYear'] == 2024].copy()

# תחזית על אותם נתונים
X_2024 = actual_2024[features]
actual_2024['PredictedQty'] = model.predict(X_2024)

# בחירת מק"ט להצגה
item_errors = []
summary_stats = []  # רשימה לאחסון נתוני סיכום המלאי

print("\n--- הערכת מודל Random Forest ומדדי מלאי עבור כל מק\"ט (2024) ---")
for item in actual_2024['ItemCode'].unique():
    try:
        item_df_2024 = actual_2024[actual_2024['ItemCode'] == item].sort_values('YearWeek')

        if len(item_df_2024) < 2:
            print(f"אין מספיק נתונים עבור מק\"ט {item} להערכת מודל או סימולציית מלאי ב-2024. מדלג.")
            continue

        # חישוב מדדי שגיאה למודל חיזוי הצריכה
        mae = mean_absolute_error(item_df_2024['ConsumptionQty'], item_df_2024['PredictedQty'])
        mse = mean_squared_error(item_df_2024['ConsumptionQty'], item_df_2024['PredictedQty'])
        rmse = np.sqrt(mse)
        r2 = r2_score(item_df_2024['ConsumptionQty'], item_df_2024['PredictedQty'])
        item_errors.append({
            'ItemCode': item,
            'MAE': mae,
            'RMSE': rmse,
            'R2': r2,
        })

        # ציור גרף צריכה חזויה מול צריכה בפועל (כפי שהיה בקוד שלך)

        plt.figure(figsize=(14, 5))
        # שינוי כאן: השתמש ב-'YearWeek' עבור ציר ה-X
        plt.plot(item_df_2024['YearWeek'], item_df_2024['ConsumptionQty'], label='Actual Consumption 2024', marker='o')
        plt.plot(item_df_2024['YearWeek'], item_df_2024['PredictedQty'], label='Forecast 2024', marker='x')

        plt.title(f'Random Forest Actual vs Predicted Consumption for ItemCode: {item} (2024)')
        plt.xlabel('Week')  # שינוי תווית ציר X
        plt.ylabel('Quantity')
        plt.xticks(rotation=45)
        plt.legend()
        plt.grid(True)
        plt.tight_layout()
        plt.show()

        # --- סימולציית מלאי (הקוד שהוספת, מותאם למודל שלך) ---
        # נתונים לסימולציה:
        # 'forecast' יהיה הצריכה החזויה של המודל
        # 'test' יהיה הצריכה בפועל

        # כדי להשתמש בפונקציית simulate_inventory, אנחנו צריכים להעביר לה DataFrames עם 'Date' ו-'Forecast'/'Actual'.
        # 'YearWeek' הוא למעשה האינדקס הזמני שלנו, אבל הפונקציה מצפה ל-'Date'.
        # ניצור עמודת תאריך מדומה מ-'YearWeek' (נניח שזה תחילת השבוע).

        # יצירת עמודת תאריך משבוע-שנה
        # הפונקציה fromisocalendar דורשת גם יום בשבוע, נניח יום שני (1)
        # נשתמש ב- ConsumptionYear וב- ConsumptionWeek מתוך item_df_2024 כדי לייצר את התאריך
        item_df_2024['SimDate'] = item_df_2024.apply(
            lambda row: datetime.fromisocalendar(int(row['ConsumptionYear']), int(row['ConsumptionWeek']), 1), axis=1
        )

        forecast_df_sim = pd.DataFrame({
            "Date": item_df_2024["SimDate"],
            "Forecast": item_df_2024["PredictedQty"]
        })
        actual_df_sim = pd.DataFrame({
            "Date": item_df_2024["SimDate"],
            "Actual": item_df_2024["ConsumptionQty"]
        })

        # הגדרות פרמטרי מלאי - אלו יכולים להיות קבועים או להגיע מהנתונים
        # לצורך הדוגמה, אנו משתמשים בערכים שהיו בקוד המקורי שלך.
        # ייתכן שתרצה להתאים את minL, maxL, initial_inventory לכל מק"ט בנפרד
        # אם יש לך נתונים כאלה.
        lt_days = 21  # 3 שבועות
        minL = 150
        maxL = 400
        initial_inventory = 300  # מלאי התחלתי לכל מק"ט בתחילת 2024

        sim_result = simulate_inventory(forecast_df_sim, actual_df_sim, lt_days, minL, maxL, initial_inventory)

        plt.figure(figsize=(12, 4))
        plt.plot(sim_result["Date"], sim_result["Inventory_End"], label="Inventory")
        plt.axhline(minL, color="green", linestyle="--", label="Min Level")
        plt.axhline(maxL, color="red", linestyle="--", label="Max Level")
        plt.title(f"Inventory Simulation for ItemCode {item} (2024)")
        plt.ylabel("Units")
        plt.xticks(rotation=45)
        plt.grid(True, linestyle='--', alpha=0.5)
        plt.legend()
        plt.tight_layout()
        plt.show()

        # ניתוח סטטוס מלאי
        analyzed_df = analyze_inventory_status(sim_result, minL, maxL)

        # אחוז בתוך תחום
        percent_within = (analyzed_df["Status"] == "Within Range").mean() * 100
        summary_stats.append({
            "ItemCode": item,
            "Percent Within Range": round(percent_within, 2)
        })

    except Exception as e:
        print(f"שגיאה במודל עבור מק\"ט {item}: {e}")

# --- סיכום כללי של מדדי המודל ---
errors_df = pd.DataFrame(item_errors)
print("\n--- סיכום מדדי ביצועי מודל RandomForest (MAE, RMSE, R2) ---")
print(errors_df.sort_values('MAE').head(25).to_string(index=False))  # מציג את 10 המק"טים עם ה-MAE הנמוך ביותר
print(errors_df.sort_values('MAE', ascending=False).head(25).to_string(
    index=False))  # מציג את 10 המק"טים עם ה-MAE הגבוה ביותר

# --- סיכום כולל של סטטוס המלאי ---
summary_df = pd.DataFrame(summary_stats)
print("\n📊 סיכום אחוזים שכל מק\"ט היה בטווח המלאי התקין (2024):")
print(summary_df.to_string(index=False))