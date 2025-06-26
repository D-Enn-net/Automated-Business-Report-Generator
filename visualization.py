# visualization.py
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os
from config import OUTPUT_REPORTS_PATH

def create_sales_visualization(metrics_data):
    """
    Creates and saves a bar chart of sales by region.
    """
    print("\nCreating sales by region visualization...")
    sales_by_region = metrics_data['sales_by_region']
    
    plt.figure(figsize=(10, 6))
    sales_by_region.plot(kind='bar', color='skyblue')
    
    plt.title('Total Sales by Region', fontsize=16)
    plt.xlabel('Region', fontsize=12)
    plt.ylabel('Total Sales ($)', fontsize=12)
    plt.xticks(rotation=0)
    plt.tight_layout()
    
    chart_path = os.path.join(OUTPUT_REPORTS_PATH, 'sales_by_region.png')
    plt.savefig(chart_path)
    plt.close()
    
    print(f"Chart saved successfully to: {chart_path}")
    return chart_path