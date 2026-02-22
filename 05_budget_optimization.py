# scripts/05_budget_optimization.py
"""
===============================================================================
 BUDGET OPTIMIZATION – LINEAR PROGRAMMING FOR CHANNEL ALLOCATION
===============================================================================
This script uses the channel performance data (from campaign_analysis.xlsx) to
optimally allocate a given total budget across channels to maximise total ROI.

It requires `pulp` – install with: pip install pulp
===============================================================================
"""

import pandas as pd
import os
import sys
import traceback

# ------------------------------------------------------------
# SETUP ERROR LOGGING
# ------------------------------------------------------------
log_file = os.path.join('..', 'output', 'budget_optimization_error.log')
try:
    import pulp
    print("=" * 60)
    print("💰 BUDGET OPTIMIZATION – CHANNEL ALLOCATION")
    print("=" * 60)

    # ------------------------------------------------------------
    # 1. LOAD CHANNEL PERFORMANCE DATA
    # ------------------------------------------------------------
    analysis_path = os.path.join('..', 'output', 'campaign_analysis.xlsx')
    if not os.path.exists(analysis_path):
        raise FileNotFoundError(f"Run 02_questions.py first to generate {analysis_path}")

    df_channels = pd.read_excel(analysis_path, sheet_name='Channel Performance')
    print("✅ Loaded channel performance data.")

    # Assume we have columns: Channel, ROI, Revenue, Spend
    channels = df_channels['Channel'].tolist()
    roi = df_channels['ROI'].values / 100  # convert percentage to decimal

    # ------------------------------------------------------------
    # 2. GET BUDGET FROM USER
    # ------------------------------------------------------------
    budget_input = input("💰 Enter total budget to allocate (e.g., 10000000 for ₹1Cr): ").strip()
    if not budget_input:
        raise ValueError("No budget entered. Exiting.")
    total_budget = float(budget_input)

    # Min and max allocation per channel (as fractions of total budget)
    min_frac = 0.05   # at least 5% per channel
    max_frac = 0.50   # at most 50% per channel

    min_spend = total_budget * min_frac
    max_spend = total_budget * max_frac

    # ------------------------------------------------------------
    # 3. BUILD LINEAR PROGRAMMING MODEL
    # ------------------------------------------------------------
    prob = pulp.LpProblem("Budget_Allocation", pulp.LpMaximize)

    # Decision variables: amount to spend on each channel
    alloc_vars = pulp.LpVariable.dicts("Alloc", channels, lowBound=0, upBound=max_spend)

    # Objective: maximise total ROI = sum(roi_i * alloc_i)
    prob += pulp.lpSum([roi[i] * alloc_vars[channels[i]] for i in range(len(channels))])

    # Constraint: total spend = total_budget
    prob += pulp.lpSum([alloc_vars[ch] for ch in channels]) == total_budget

    # Min allocation per channel
    for ch in channels:
        prob += alloc_vars[ch] >= min_spend

    # ------------------------------------------------------------
    # 4. SOLVE
    # ------------------------------------------------------------
    prob.solve(pulp.PULP_CBC_CMD(msg=0))

    # ------------------------------------------------------------
    # 5. DISPLAY RESULTS
    # ------------------------------------------------------------
    print("\n📊 Optimal Allocation:")
    allocation = []
    total_roi = 0
    for i, ch in enumerate(channels):
        amt = alloc_vars[ch].varValue
        if amt is None:
            raise RuntimeError(f"No solution found for channel {ch}")
        allocation.append({
            'Channel': ch,
            'Allocated Budget (₹)': round(amt, 0),
            'Expected ROI %': round(roi[i] * 100, 2)
        })
        total_roi += roi[i] * amt

    df_alloc = pd.DataFrame(allocation)
    df_alloc['Allocated Budget (₹)'] = df_alloc['Allocated Budget (₹)'].astype(int)
    print(df_alloc.to_string(index=False))
    print(f"\n💰 Total Expected ROI Value: ₹{total_roi:,.0f}")

    # ------------------------------------------------------------
    # 6. SAVE TO EXCEL
    # ------------------------------------------------------------
    output_path = os.path.join('..', 'output', 'budget_optimization.xlsx')
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_alloc.to_excel(writer, sheet_name='Optimal Allocation', index=False)
        df_channels.to_excel(writer, sheet_name='Channel Performance', index=False)

    print(f"✅ Results saved to:\n   {output_path}")

except Exception as e:
    print("\n❌ AN ERROR OCCURRED:")
    print(traceback.format_exc())
    with open(log_file, 'w') as f:
        f.write(traceback.format_exc())
    print(f"\nError details written to: {log_file}")
finally:
    input("\nPress Enter to exit...")