# scripts/04_audience_segmentation.py
"""
===============================================================================
 AUDIENCE SEGMENTATION – CLUSTER TARGET AUDIENCES BY PERFORMANCE
===============================================================================
This script aggregates campaign data by target audience, then applies K‑Means
clustering to segment audiences based on:
   • Total revenue
   • Average ROI
   • Average CTR
   • Average CVR
   • Total conversions
   • Number of campaigns

Output: output/audience_segments.xlsx (sheets: Audience Profiles, Audience Clusters)
===============================================================================
"""

import pandas as pd
import numpy as np
import os
import sqlite3
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA

print("=" * 60)
print("👥 AUDIENCE SEGMENTATION – CLUSTERING TARGET AUDIENCES")
print("=" * 60)

# ------------------------------------------------------------
# 1. LOAD CLEANED DATA
# ------------------------------------------------------------
db_path = os.path.join('..', 'output', 'marketing.db')
excel_path = os.path.join('..', 'output', 'cleaned_marketing_data.xlsx')
output_path = os.path.join('..', 'output', 'audience_segments.xlsx')

if os.path.exists(db_path):
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT * FROM campaigns", conn)
    conn.close()
    print("✅ Loaded from SQLite.")
elif os.path.exists(excel_path):
    df = pd.read_excel(excel_path)
    print("✅ Loaded from Excel.")
else:
    raise FileNotFoundError("No cleaned data found.")

print(f"📊 Total campaigns: {len(df):,}")

# ------------------------------------------------------------
# 2. AGGREGATE BY TARGET AUDIENCE
# ------------------------------------------------------------
audience_agg = df.groupby('Target_Audience').agg({
    'Revenue': 'sum',
    'Spend': 'sum',
    'Conversion': 'sum',
    'Clicks': 'sum',
    'Impressions': 'sum',
    'ROI': 'mean',
    'CTR': 'mean',
    'CVR': 'mean',
    'CPA': 'mean',
    'Campaign_Id': 'count'
}).rename(columns={'Campaign_Id': 'Campaign_Count'}).reset_index()

# Clean up infinite/NaN
for col in ['ROI', 'CTR', 'CVR', 'CPA']:
    audience_agg[col] = audience_agg[col].replace([np.inf, -np.inf], np.nan).fillna(0)

print(f"✅ Aggregated {len(audience_agg)} target audiences.")

# ------------------------------------------------------------
# 3. SELECT FEATURES FOR CLUSTERING
# ------------------------------------------------------------
feature_cols = ['Revenue', 'ROI', 'CTR', 'CVR', 'Conversion', 'Campaign_Count']
X = audience_agg[feature_cols].fillna(0).values

# Standardise
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

# ------------------------------------------------------------
# 4. FIND OPTIMAL K (Elbow) – simple, use K=3
# ------------------------------------------------------------
# For simplicity, use K=3 (you can adjust)
k_audience = 3
kmeans = KMeans(n_clusters=k_audience, random_state=42, n_init=10)
audience_agg['Cluster'] = kmeans.fit_predict(X_scaled)

# ------------------------------------------------------------
# 5. INTERPRET CLUSTERS
# ------------------------------------------------------------
cluster_profile = audience_agg.groupby('Cluster')[feature_cols].mean().round(2)
cluster_profile['Size'] = audience_agg.groupby('Cluster').size()

# Label clusters
def label_audience_cluster(row):
    if row['Revenue'] > cluster_profile['Revenue'].median() and row['ROI'] > cluster_profile['ROI'].median():
        return "High Value"
    elif row['Revenue'] < cluster_profile['Revenue'].median() and row['ROI'] < 0:
        return "Low Value / Loss"
    else:
        return "Mid Tier"

cluster_profile['Label'] = cluster_profile.apply(label_audience_cluster, axis=1)

# Map labels back
label_map = cluster_profile['Label'].to_dict()
audience_agg['Cluster_Label'] = audience_agg['Cluster'].map(label_map)

print("\n📊 Audience Cluster Profiles:")
print(cluster_profile)

# ------------------------------------------------------------
# 6. SAVE TO EXCEL
# ------------------------------------------------------------
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    audience_agg.to_excel(writer, sheet_name='Audience Profiles', index=False)
    cluster_profile.reset_index().to_excel(writer, sheet_name='Cluster Summary', index=False)

print(f"✅ Audience segments saved to:\n   {output_path}")