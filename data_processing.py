# data_processing.py
import pandas as pd

def analyze_sales_data(path_to_file):
    """
    Loads and analyzes sales data to extract key business metrics.
    """
    print(f"Attempting to load file from: {path_to_file}")
    try:
        df = pd.read_csv(path_to_file, encoding='latin1')
        print("File loaded successfully!")
        
        print("\nCleaning data...")
        df['Order Date'] = pd.to_datetime(df['Order Date'], dayfirst=True)
        df['Month'] = df['Order Date'].dt.to_period('M')
        print("Data cleaning complete.")

        print("\nCalculating key metrics...")
        total_sales = df['Sales'].sum()
        sales_by_region = df.groupby('Region')['Sales'].sum().sort_values(ascending=False)
        sales_by_category = df.groupby('Category')['Sales'].sum().sort_values(ascending=False)
        print("Metrics calculated successfully!")
        
        return {
            'total_sales': total_sales,
            'sales_by_region': sales_by_region,
            'sales_by_category': sales_by_category
        }
    except Exception as e:
        print(f"An unexpected error occurred during analysis: {e}")
        return None