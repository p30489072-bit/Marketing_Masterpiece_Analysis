# scripts/03_clustering.py
"""
===============================================================================
 MARKETING CLUSTERING – GROUP CAMPAIGNS BY PERFORMANCE
===============================================================================
This script loads cleaned campaign data and performs:
   • K‑Means clustering on performance metrics (CTR, CVR, ROI, CPA, Revenue per impression)
   • Feature importance (Random Forest) to identify what drives revenue
   • PCA for 2D visualisation (optional)

Output: output/ml_insights.xlsx (sheets: Campaign Clusters, Cluster Profiles, Feature Importance, PCA Coordinates)
===============================================================================
"""

import pandas as pd
import numpy as np
import os
import sqlite3
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from sklearn.ensemble import RandomForestRegressor
from sklearn.decomposition import PCA
import warnings
warnings.filterwarnings('ignore')

# ------------------------------------------------------------
# 1. LOAD CLEANED DATA
# ------------------------------------------------------------
print("=" * 60)
print("📊 MARKETING CLUSTERING – GROUPING CAMPAIGNS")
print("=" * 60)

db_path = os.path.join('..', 'output', 'marketing.db')
excel_path = os.path.join('..', 'output', 'cleaned_marketing_data.xlsx')
output_path = os.path.join('..', 'output', 'ml_insights.xlsx')

if os.path.exists(db_path):
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT * FROM campaigns", conn)
    conn.close()
    print("✅ Loaded from SQLite.")
elif os.path.exists(excel_path):
    df = pd.read_excel(excel_path)
    print("✅ Loaded from Excel.")
else:
    raise FileNotFoundError("No cleaned data found in output folder.")

print(f"📊 Total campaigns: {len(df):,}")

# ------------------------------------------------------------
# 2. PREPARE FEATURES FOR CLUSTERING
# ------------------------------------------------------------
feature_cols = ['CTR', 'CVR', 'ROI', 'CPA', 'Revenue_Per_Impression']

# Ensure numeric and drop rows with missing/infinite
for col in feature_cols:
    df[col] = pd.to_numeric(df[col], errors='coerce')
df = df.replace([np.inf, -np.inf], np.nan).dropna(subset=feature_cols).copy()

X = df[feature_cols].values

# Standardise features
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

print(f"✅ Features prepared: {feature_cols}")

# ------------------------------------------------------------
# 3. FIND OPTIMAL NUMBER OF CLUSTERS (Elbow method)
# ------------------------------------------------------------
print("🔍 Finding optimal clusters (elbow method)...")
inertia = []
K_range = range(2, 10)
for k in K_range:
    kmeans = KMeans(n_clusters=k, random_state=42, n_init=10)
    kmeans.fit(X_scaled)
    inertia.append(kmeans.inertia_)

# (Optional) print elbow values – you can plot later if needed
print("   Inertia for K=2..9:", [round(i,0) for i in inertia])

# Choose K (we'll use 4 – you can adjust based on elbow)
k_optimal = 4
print(f"✅ Using K = {k_optimal} clusters (adjustable in script).")

# ------------------------------------------------------------
# 4. APPLY K‑MEANS
# ------------------------------------------------------------
kmeans = KMeans(n_clusters=k_optimal, random_state=42, n_init=10)
df['Cluster'] = kmeans.fit_predict(X_scaled)

# ------------------------------------------------------------
# 5. INTERPRET CLUSTERS – CREATE PROFILES
# ------------------------------------------------------------
cluster_profile = df.groupby('Cluster')[feature_cols + ['Revenue', 'Spend']].mean().round(2)
cluster_profile['Size'] = df.groupby('Cluster').size()

# Add simple labels based on ROI
def label_cluster(row):
    if row['ROI'] > cluster_profile['ROI'].median():
        return "High ROI"
    elif row['ROI'] < 0:
        return "Loss Maker"
    elif row['CPA'] > cluster_profile['CPA'].median():
        return "High Cost"
    else:
        return "Average"

cluster_profile['Label'] = cluster_profile.apply(label_cluster, axis=1)

# Map labels back to campaigns
label_map = cluster_profile['Label'].to_dict()
df['Cluster_Label'] = df['Cluster'].map(label_map)

print("\n📊 Cluster Profiles:")
print(cluster_profile)

# ------------------------------------------------------------
# 6. FEATURE IMPORTANCE (Random Forest)
# ------------------------------------------------------------
print("\n🔍 Analysing feature importance (predicting Revenue)...")
rf = RandomForestRegressor(n_estimators=100, random_state=42)
rf.fit(X_scaled, df['Revenue'])
importance = pd.DataFrame({
    'Feature': feature_cols,
    'Importance': rf.feature_importances_
}).sort_values('Importance', ascending=False)
print(importance)

# ------------------------------------------------------------
# 7. PCA FOR 2D VISUALISATION (optional)
# ------------------------------------------------------------
pca = PCA(n_components=2)
X_pca = pca.fit_transform(X_scaled)
pca_df = pd.DataFrame(X_pca, columns=['PC1', 'PC2'])
pca_df['Cluster'] = df['Cluster']
pca_df['Cluster_Label'] = df['Cluster_Label']

# ------------------------------------------------------------
# 8. SAVE TO EXCEL
# ------------------------------------------------------------
print("\n💾 Saving results to Excel...")

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    # Campaigns with cluster labels
    df[['Campaign_Id', 'Campaign_Name', 'Channel', 'CTR', 'CVR', 'ROI', 'CPA', 'Revenue_Per_Impression', 'Cluster', 'Cluster_Label']].to_excel(writer, sheet_name='Campaign Clusters', index=False)
    # Cluster profiles
    cluster_profile.reset_index().to_excel(writer, sheet_name='Cluster Profiles', index=False)
    # Feature importance
    importance.to_excel(writer, sheet_name='Feature Importance', index=False)
    # PCA coordinates (optional)
    pca_df.to_excel(writer, sheet_name='PCA Coordinates', index=False)

print(f"✅ ML insights saved to:\n   {output_path}")

# ------------------------------------------------------------
# 9. CONSOLE SUMMARY
# ------------------------------------------------------------
print("\n" + "="*60)
print("📊 CLUSTERING SUMMARY")
print("="*60)
for i in range(k_optimal):
    label = cluster_profile.loc[i, 'Label']
    size = cluster_profile.loc[i, 'Size']
    roi = cluster_profile.loc[i, 'ROI']
    ctr = cluster_profile.loc[i, 'CTR']
    print(f"\nCluster {i}: {label} ({size} campaigns)")
    print(f"   Avg ROI: {roi:.1f}% | Avg CTR: {ctr:.2f}%")
print("\n✅ Top features driving revenue:")
for _, row in importance.iterrows():
    print(f"   {row['Feature']}: {row['Importance']:.3f}")