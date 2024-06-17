import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = 'Include/mar_2023.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Extract relevant columns
relevant_columns = [
    'Company Name',
    'Total assets',
    'Total liabilities',
    'Net working capital',
    'Gross working capital (cost of sales method)',
    'Gross working capital cycle (days)',
    'Net working capital cycle (days)',
    'Short term loans & advances by finance companies to current assets',
    'Raw material cycle (days)',
    'Debtor days (days)',
    'Short term investments to current assets',
    'Trade receivables & bills receivables to current assets',
    'Short term loans & advances to current assets',
    'Cash & bank balance to current assets'
]

# Filter the dataframe to only include relevant columns
filtered_df = df[relevant_columns]

# Calculate ratios
filtered_df['Debt to Equity Ratio'] = filtered_df['Total liabilities'] / filtered_df['Total assets']
filtered_df['Interest Coverage Ratio'] = filtered_df['Gross working capital (cost of sales method)'] / filtered_df['Short term loans & advances by finance companies to current assets']
filtered_df['Current Ratio'] = filtered_df['Total assets'] / filtered_df['Total liabilities']
filtered_df['Cash Ratio'] = filtered_df['Cash & bank balance to current assets'] / filtered_df['Total liabilities']
filtered_df['Average Age of Inventory'] = (filtered_df['Net working capital'] / filtered_df['Gross working capital (cost of sales method)']) * 365
filtered_df['Working Capital Turnover Ratio'] = filtered_df['Gross working capital (cost of sales method)'] / filtered_df['Net working capital']
filtered_df['Asset Turnover Ratio'] = filtered_df['Gross working capital (cost of sales method)'] / filtered_df['Total assets']
filtered_df['Fixed Asset Turnover Ratio'] = filtered_df['Gross working capital (cost of sales method)'] / filtered_df['Net working capital']

# Plotting

plt.figure(figsize=(14, 10))

# Working Capital Turnover Ratio Distribution
sns.histplot(filtered_df['Working Capital Turnover Ratio'], kde=True)
plt.title('Working Capital Turnover Ratio Distribution')
plt.show()

