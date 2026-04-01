import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec

# --- Data load & clean ---
df = pd.read_csv(r"C:\Users\gyans\OneDrive\Documents\New folder\Sample - Superstore.csv", encoding='latin1' )
df['Order Date'] = pd.to_datetime(df['Order Date'])
df['Month'] = df['Order Date'].dt.to_period('M')

# --- Calculations ---
monthly_sales = df.groupby('Month')['Sales'].sum()
category_profit = df.groupby('Category')['Profit'].sum()
region_sales = df.groupby('Region')['Sales'].sum()
top_products = df.groupby('Sub-Category')['Sales'].sum().nlargest(5)

# --- Dashboard layout ---
fig = plt.figure(figsize=(14, 10))
fig.suptitle('Sales Dashboard — Superstore', fontsize=16, fontweight='bold', y=0.98)
gs = gridspec.GridSpec(2, 2, figure=fig, hspace=0.4, wspace=0.35)

# Chart 1: Monthly Sales Trend
ax1 = fig.add_subplot(gs[0, :])
ax1.plot(monthly_sales.index.astype(str), monthly_sales.values,
         color='#378ADD', linewidth=2, marker='o', markersize=4)
ax1.set_title('Monthly Sales Trend', fontweight='bold')
ax1.set_xlabel('Month')
ax1.set_ylabel('Sales (₹)')
ax1.tick_params(axis='x', rotation=45)
ax1.grid(axis='y', alpha=0.3)

# Chart 2: Profit by Category
ax2 = fig.add_subplot(gs[1, 0])
colors = ['#1D9E75' if v > 0 else '#E24B4A' for v in category_profit.values]
ax2.bar(category_profit.index, category_profit.values, color=colors)
ax2.set_title('Profit by Category', fontweight='bold')
ax2.set_ylabel('Profit (₹)')
ax2.grid(axis='y', alpha=0.3)

# Chart 3: Region-wise Sales (Pie)
ax3 = fig.add_subplot(gs[1, 1])
ax3.pie(region_sales.values, labels=region_sales.index,
        autopct='%1.1f%%', colors=['#378ADD','#1D9E75','#EF9F27','#D4537E'])
ax3.set_title('Sales by Region', fontweight='bold')

import os
os.makedirs('output', exist_ok=True)
plt.savefig('output/dashboard.png', dpi=150, bbox_inches='tight')
print("Dashboard saved!")
plt.show()