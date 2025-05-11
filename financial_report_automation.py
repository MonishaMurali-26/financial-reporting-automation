import pandas as pd
import numpy as np
import sqlite3
import os
import matplotlib.pyplot as plt
import seaborn as sns

def generate_financial_data():
    np.random.seed(42)
    num_months = 24
    data = {
        'month': pd.date_range(start='2022-01-01', periods=num_months, freq='M'),
        'revenue': np.random.normal(75000, 10000, num_months),
        'expenses': np.random.normal(45000, 8000, num_months),
        'sales_volume': np.random.poisson(500, num_months)
    }
    df = pd.DataFrame(data)
    df['revenue'] = df['revenue'].clip(lower=20000).round(2)
    df['expenses'] = df['expenses'].clip(lower=15000).round(2)
    df['sales_volume'] = df['sales_volume'].clip(lower=100)
    if not os.path.exists("data"):
        os.makedirs("data")
    df.to_csv('data/financial_data.csv', index=False)
    return df

def store_in_warehouse(df):
    conn = sqlite3.connect("financial_warehouse.db")
    df.to_sql('financials', conn, if_exists='replace', index=False)
    conn.close()

def query_financial_data():
    conn = sqlite3.connect("financial_warehouse.db")
    query = """
    SELECT 
        month,
        revenue,
        expenses,
        sales_volume,
        (revenue - expenses) AS profit,
        ((revenue - expenses) / revenue * 100) AS profit_margin,
        (revenue / sales_volume) AS revenue_per_unit
    FROM financials
    ORDER BY month
    """
    df = pd.read_sql_query(query, conn)
    conn.close()
    df['month'] = pd.to_datetime(df['month'])
    return df

def create_visualizations(df):
    plt.figure(figsize=(10, 6))
    plt.plot(df['month'], df['revenue'], label='Revenue', color='blue')
    plt.plot(df['month'], df['expenses'], label='Expenses', color='red')
    plt.title('Revenue vs Expenses Over Time')
    plt.xlabel('Month')
    plt.ylabel('Amount ($)')
    plt.legend()
    plt.grid(True)
    plt.savefig('revenue_vs_expenses.png')
    plt.close()

    plt.figure(figsize=(10, 6))
    plt.fill_between(df['month'], df['profit_margin'], color='green', alpha=0.5)
    plt.title('Profit Margin Trend')
    plt.xlabel('Month')
    plt.ylabel('Profit Margin (%)')
    plt.grid(True)
    plt.savefig('profit_margin_trend.png')
    plt.close()

    plt.figure(figsize=(10, 6))
    sns.histplot(df['revenue_per_unit'], bins=20, kde=True, color='purple')
    plt.title('Revenue per Unit Distribution')
    plt.xlabel('Revenue per Unit ($)')
    plt.ylabel('Frequency')
    plt.savefig('revenue_per_unit_histogram.png')
    plt.close()

    plt.figure(figsize=(10, 6))
    plt.bar(df['month'], df['profit'], color='orange')
    plt.title('Profit by Month')
    plt.xlabel('Month')
    plt.ylabel('Profit ($)')
    plt.xticks(rotation=45)
    plt.grid(True, axis='y')
    plt.savefig('profit_by_month.png')
    plt.close()

def export_to_excel(df):
    df['revenue'] = df['revenue'].round(2)
    df['expenses'] = df['expenses'].round(2)
    df['profit'] = df['profit'].round(2)
    df['profit_margin'] = df['profit_margin'].round(2)
    df['revenue_per_unit'] = df['revenue_per_unit'].round(2)
    df.to_excel('financial_report.xlsx', index=False, sheet_name='Monthly Financial Report')
    print("Financial report generated and saved to 'financial_report.xlsx'.")

if __name__ == "__main__":
    print("Starting Financial Reporting Automation Pipeline...")
    financial_df = generate_financial_data()
    store_in_warehouse(financial_df)
    report_df = query_financial_data()
    create_visualizations(report_df)
    export_to_excel(report_df)
    print("Pipeline completed successfully!")