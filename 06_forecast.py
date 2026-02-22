# scripts/06_forecast.py
"""
===============================================================================
 REVENUE FORECASTING WITH PROPHET – FIXED COLUMN DETECTION
===============================================================================
"""

import pandas as pd
import numpy as np
import os
import sqlite3
from prophet import Prophet
import matplotlib.pyplot as plt

print("=" * 60)
print("📈 REVENUE FORECASTING WITH PROPHET")
print("=" * 60)

# ------------------------------------------------------------
# 1. LOAD CLEANED DATA
# ------------------------------------------------------------
db_path = os.path.join('..', 'output', 'marketing.db')
excel_path = os.path.join('..', 'output', 'cleaned_marketing_data.xlsx')

if os.path.exists(db_path):
    conn = sqlite3.connect(db_path)
    # Read all columns; we'll find date and revenue later
    df = pd.read_sql_query("SELECT * FROM campaigns", conn)
    conn.close()
    print("✅ Loaded from SQLite.")
elif os.path.exists(excel_path):
    df = pd.read_excel(excel_path)
    print("✅ Loaded from Excel.")
else:
    raise FileNotFoundError("No cleaned data found.")

# ------------------------------------------------------------
# 2. FIND DATE AND REVENUE COLUMNS (case‑insensitive)
# ------------------------------------------------------------
date_col = None
revenue_col = None
for col in df.columns:
    if col.lower() == 'date':
        date_col = col
    elif col.lower() == 'revenue':
        revenue_col = col

if date_col is None:
    raise KeyError("No column named 'Date' or 'date' found in data.")
if revenue_col is None:
    raise KeyError("No column named 'Revenue' or 'revenue' found in data.")

print(f"📅 Using date column: {date_col}")
print(f"💰 Using revenue column: {revenue_col}")

# Ensure date is datetime
df[date_col] = pd.to_datetime(df[date_col])

# Aggregate daily revenue (in case multiple campaigns per day)
daily = df.groupby(df[date_col].dt.date)[revenue_col].sum().reset_index()
daily.columns = ['ds', 'y']
daily['ds'] = pd.to_datetime(daily['ds'])
daily = daily.sort_values('ds')

print(f"📊 Daily revenue data: {len(daily)} days")

# ------------------------------------------------------------
# 3. FIT PROPHET MODEL
# ------------------------------------------------------------
model = Prophet(yearly_seasonality=True, weekly_seasonality=True, daily_seasonality=False)
model.fit(daily)

# Create future dataframe for next 90 days
future = model.make_future_dataframe(periods=90)
forecast = model.predict(future)

# ------------------------------------------------------------
# 4. PLOT FORECAST
# ------------------------------------------------------------
fig = model.plot(forecast)
plt.title('Revenue Forecast (Next 90 Days)')
plt.xlabel('Date')
plt.ylabel('Revenue (₹)')
plot_path = os.path.join('..', 'output', 'forecast_plot.png')
plt.savefig(plot_path, dpi=150, bbox_inches='tight')
plt.close()
print(f"✅ Forecast plot saved to: {plot_path}")

# ------------------------------------------------------------
# 5. SAVE FORECAST DATA TO EXCEL
# ------------------------------------------------------------
output_path = os.path.join('..', 'output', 'forecast.xlsx')
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].tail(100).to_excel(writer, sheet_name='Forecast', index=False)
    # Also save the components
    model.plot_components(forecast)
    plt.savefig(os.path.join('..', 'output', 'forecast_components.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("✅ Forecast components plot saved.")

print(f"✅ Forecast data saved to: {output_path}")

# ------------------------------------------------------------
# 6. PRINT KEY FORECAST INSIGHTS
# ------------------------------------------------------------
latest = forecast.iloc[-1]
print("\n🔮 Key Forecast Insights")
print("-" * 40)
print(f"Last historical date: {daily['ds'].max().date()}")
print(f"Forecast horizon: 90 days")
print(f"Predicted total revenue next 90 days: ₹{forecast['yhat'].tail(90).sum():,.0f}")
print(f"Average daily revenue next 90 days: ₹{forecast['yhat'].tail(90).mean():,.0f}")

# ------------------------------------------------------------
# 7. PAUSE TO KEEP WINDOW OPEN
# ------------------------------------------------------------
input("\n✅ Forecast complete. Press Enter to exit...")