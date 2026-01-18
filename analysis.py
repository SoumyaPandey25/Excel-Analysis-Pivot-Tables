# =====================================
# TASK 2 – Pivot Table Analysis
# Using Superstore Sales Dataset
# =====================================

import pandas as pd

# -------------------------------------
# STEP 1: Load Dataset
# -------------------------------------
df = pd.read_csv("Superstore_Sales_Dataset.csv")

# Convert Order Date to datetime
df["Order Date"] = pd.to_datetime(df["Order Date"], dayfirst=True)

print("\nDataset Loaded Successfully!")
print(df.head())

# -------------------------------------
# STEP 2: Pivot 1 – Total Sales by Category
# -------------------------------------
pivot_sales_category = pd.pivot_table(
    df,
    values="Sales",
    index="Category",
    aggfunc="sum"
).sort_values(by="Sales", ascending=False)

print("\nTotal Sales by Category:")
print(pivot_sales_category)

# -------------------------------------
# STEP 3: Pivot 2 – Sales by Region and Segment
# -------------------------------------
pivot_sales_region_segment = pd.pivot_table(
    df,
    values="Sales",
    index="Region",
    columns="Segment",
    aggfunc="sum"
)

print("\nSales by Region and Segment:")
print(pivot_sales_region_segment)

# -------------------------------------
# STEP 4: Pivot 3 – Sales by Sub-Category
# -------------------------------------
pivot_sales_subcategory = pd.pivot_table(
    df,
    values="Sales",
    index="Sub-Category",
    aggfunc="sum"
).sort_values(by="Sales", ascending=False)

print("\nSales by Sub-Category:")
print(pivot_sales_subcategory)

# -------------------------------------
# STEP 5: Identify Top & Underperforming Regions
# -------------------------------------
region_sales = df.groupby("Region")["Sales"].sum().sort_values()

print("\nRegion-wise Sales:")
print(region_sales)

print("\nUnderperforming Region:", region_sales.idxmin())
print("Top Performing Region:", region_sales.idxmax())

# -------------------------------------
# STEP 6: Save Pivot Tables to Excel
# -------------------------------------
with pd.ExcelWriter("Pivot_Report.xlsx") as writer:
    pivot_sales_category.to_excel(writer, sheet_name="Sales_by_Category")
    pivot_sales_region_segment.to_excel(writer, sheet_name="Sales_by_Region_Segment")
    pivot_sales_subcategory.to_excel(writer, sheet_name="Sales_by_SubCategory")

print("\nPivot_Report.xlsx created successfully!")

# -------------------------------------
# STEP 7: Write Insights
# -------------------------------------
insights = [
    "Technology category contributes the highest total sales.",
    "West region is the top-performing region based on total sales.",
    "Consumer segment generates more sales compared to Corporate and Home Office.",
    "Some sub-categories contribute disproportionately to overall sales.",
    "Certain regions are underperforming and need focused business strategies."
]

with open("Insights.txt", "w") as f:
    for i, insight in enumerate(insights, 1):
        f.write(f"{i}. {insight}\n")

print("Insights.txt created successfully!")
