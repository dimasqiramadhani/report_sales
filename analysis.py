import pandas as pd
import matplotlib.pyplot as plt
import os

# Set the path to the data file
file_path = 'data/sales_data.xlsx'

# Load data from Excel
try:
    data = pd.read_excel(file_path)
except FileNotFoundError as e:
    print(f"Error: {e}")
    exit()

# Data overview
print("Data Overview:")
print(data.head())

# Total sales per product
total_sales_per_product = data.groupby('Product')['Total Sales'].sum().reset_index()
print("\nTotal Sales Per Product:")
print(total_sales_per_product)

# Total sales over time
total_sales_over_time = data.groupby('Date')['Total Sales'].sum().reset_index()
print("\nTotal Sales Over Time:")
print(total_sales_over_time)

# Save analysis results to CSV
total_sales_per_product.to_csv('reports/total_sales_per_product.csv', index=False)
total_sales_over_time.to_csv('reports/total_sales_over_time.csv', index=False)

# Plot total sales per product
plt.figure(figsize=(10, 6))
plt.bar(total_sales_per_product['Product'], total_sales_per_product['Total Sales'], color='skyblue')
plt.xlabel('Product')
plt.ylabel('Total Sales')
plt.title('Total Sales Per Product')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('reports/total_sales_per_product.png')
plt.show()

# Plot total sales over time
plt.figure(figsize=(10, 6))
plt.plot(total_sales_over_time['Date'], total_sales_over_time['Total Sales'], marker='o', linestyle='-', color='orange')
plt.xlabel('Date')
plt.ylabel('Total Sales')
plt.title('Total Sales Over Time')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('reports/sales_trend.png')
plt.show()
