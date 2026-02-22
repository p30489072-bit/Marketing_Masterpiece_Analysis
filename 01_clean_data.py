# scripts/02_clean_data.py
import pandas as pd
import numpy as np
import os
import sqlite3
from datetime import datetime

# Paths
data_folder = os.path.join('..', 'data')
output_folder = os.path.join('..', 'output')
db_path = os.path.join(output_folder, 'marketing.db')
excel_path = os.path.join(output_folder, 'cleaned_marketing_data.xlsx')

# Ensure output folder exists
os.makedirs(output_folder, exist_ok=True)

# Find the Excel file
files = os.listdir(data_folder)
excel_files = [f for f in files if f.endswith(('.xlsx', '.xls'))]
if not excel_files:
    raise FileNotFoundError("No Excel file found in data folder.")

data_path = os.path.join(data_folder, excel_files[0])
print(f"📂 Loading: {data_path}")
df = pd.read_excel(data_path)
print(f"✅ Loaded {len(df)} rows.")

# ------------------------------------------------------------
# 1. Standardize column names (internal use only)
# ------------------------------------------------------------
df.columns = [c.strip().lower().replace(' ', '_') for c in df.columns]

# ------------------------------------------------------------
# 2. Ensure date column is datetime
# ------------------------------------------------------------
df['date'] = pd.to_datetime(df['date'], errors='coerce')
initial_len = len(df)
df = df.dropna(subset=['date'])
if len(df) < initial_len:
    print(f"⚠️ Dropped {initial_len - len(df)} rows with invalid dates.")

# ------------------------------------------------------------
# 3. Convert numeric columns
# ------------------------------------------------------------
numeric_cols = ['impressions', 'clicks', 'conversions', 'revenue', 'spend']
for col in numeric_cols:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# ------------------------------------------------------------
# 4. Add derived marketing metrics
# ------------------------------------------------------------
df['ctr'] = (df['clicks'] / df['impressions']) * 100
df['ctr'] = df['ctr'].replace([np.inf, -np.inf], np.nan).fillna(0)

df['cvr'] = (df['conversions'] / df['clicks']) * 100
df['cvr'] = df['cvr'].replace([np.inf, -np.inf], np.nan).fillna(0)

df['cpa'] = df['spend'] / df['conversions']
df['cpa'] = df['cpa'].replace([np.inf, -np.inf], np.nan).fillna(0)

df['roi'] = (df['revenue'] - df['spend']) / df['spend'] * 100
df['roi'] = df['roi'].replace([np.inf, -np.inf], np.nan).fillna(0)

df['revenue_per_impression'] = df['revenue'] / df['impressions']
df['revenue_per_impression'] = df['revenue_per_impression'].replace([np.inf, -np.inf], np.nan).fillna(0)

# ------------------------------------------------------------
# 5. Rename columns to YOUR headers (exactly as you want)
# ------------------------------------------------------------
df.rename(columns={
    'campaign_id': 'Campaign_Id',
    'campaign_name': 'Campaign_Name',
    'channel': 'Channel',
    'date': 'Date',
    'impressions': 'Impressions',
    'clicks': 'Clicks',
    'conversions': 'Conversion',
    'revenue': 'Revenue',
    'spend': 'Spend',
    'region': 'Region',
    'target_audience': 'Target_Audience',
    'creative_type': 'Creative_Type',
    'ctr': 'CTR',
    'cvr': 'CVR',
    'cpa': 'CPA',
    'roi': 'ROI',
    'revenue_per_impression': 'Revenue_Per_Impression'
}, inplace=True)

# ------------------------------------------------------------
# 6. Standardize categorical columns
# ------------------------------------------------------------
categorical_cols = ['Channel', 'Region', 'Target_Audience', 'Creative_Type']
for col in categorical_cols:
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip().str.title()

# ------------------------------------------------------------
# 7. Save to SQLite
# ------------------------------------------------------------
conn = sqlite3.connect(db_path)
df.to_sql('campaigns', conn, if_exists='replace', index=False)
print(f"💾 Saved to SQLite: {db_path}")

# ------------------------------------------------------------
# 8. Save to Excel with YOUR headers and proper formatting
# ------------------------------------------------------------
with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Cleaned_Data', index=False)
    
    # Get workbook and worksheet
    workbook = writer.book
    worksheet = writer.sheets['Cleaned_Data']
    
    # Apply percentage formatting to CTR, CVR, ROI columns
    for col_idx, col_name in enumerate(df.columns, start=1):
        if col_name in ['CTR', 'CVR', 'ROI']:
            # Apply percentage format with 2 decimal places
            for row in range(2, len(df) + 2):
                cell = worksheet.cell(row=row, column=col_idx)
                cell.number_format = '0.00%'
        elif col_name in ['CPA', 'Revenue_Per_Impression']:
            # Currency format for CPA and revenue per impression
            for row in range(2, len(df) + 2):
                cell = worksheet.cell(row=row, column=col_idx)
                cell.number_format = '₹#,##0.00'

print(f"💾 Saved to Excel with percentage formatting: {excel_path}")

# ------------------------------------------------------------
# 9. Quick summary
# ------------------------------------------------------------
print("\n📊 Cleaned Data Summary")
print("─" * 40)
print(f"Total campaigns: {len(df)}")
print(f"Date range: {df['Date'].min().date()} to {df['Date'].max().date()}")
print(f"Total spend: ₹{df['Spend'].sum():,.0f}")
print(f"Total revenue: ₹{df['Revenue'].sum():,.0f}")
print(f"Average ROI: {df['ROI'].mean():.1f}%")
print(f"Average CTR: {df['CTR'].mean():.2f}%")
print(f"Average CVR: {df['CVR'].mean():.2f}%")
print(f"Average CPA: ₹{df['CPA'].mean():,.0f}")

print("\n✅ Cleaning complete! Next step: analysis and dashboard.")