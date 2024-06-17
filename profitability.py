import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = 'Include/mar_2023.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Display the first few rows of the dataframe to understand its structure
print(df.head())
print(df.columns)

# Extract relevant columns
relevant_columns = [
    'Company Name',
    'Total assets',
    'Total liabilities',
    'Net working capital',
    'Gross working capital (cost of sales method)',
    'Gross working capital cycle (days)',
    'Net working capital cycle (days)'
]

# Filter the dataframe to only include relevant columns
filtered_df = df[relevant_columns]

# Display the first few rows of the filtered dataframe
print(filtered_df.head())
print(filtered_df.columns)

# Calculate the required ratios

# Asset Turnover Ratio = Revenue / Total Assets
filtered_df['Asset Turnover Ratio'] = filtered_df['Gross working capital (cost of sales method)'] / filtered_df['Total assets']

# Fixed Asset Turnover Ratio = Revenue / Fixed Assets
filtered_df['Fixed Asset Turnover Ratio'] = filtered_df['Gross working capital (cost of sales method)'] / filtered_df['Total assets']

# Working Capital Turnover Ratio = Revenue / Net Working Capital
filtered_df['Working Capital Turnover Ratio'] = filtered_df['Gross working capital (cost of sales method)'] / filtered_df['Net working capital']

# Display the first few rows of the dataframe with the new ratios
print(filtered_df.head())
print('done')


# Set the figure size
plt.figure(figsize=(14, 8))

# Create a bar plot for Asset Turnover Ratio
sns.barplot(x='Company Name', y='Asset Turnover Ratio', data=filtered_df, palette='magma')

# Rotate the x labels for better readability
plt.xticks(rotation=90)

# Set the title and labels
plt.title('Asset Turnover Ratio for All Companies')
plt.xlabel('Company Name')
plt.ylabel('Asset Turnover Ratio')

# Show the plot
plt.show()


# Set the figure size
plt.figure(figsize=(14, 8))

# Create a bar plot for Asset Turnover Ratio
sns.barplot(x='Company Name', y='Fixed Asset Turnover Ratio', data=filtered_df, palette='magma')

# Rotate the x labels for better readability
plt.xticks(rotation=90)

# Set the title and labels
plt.title('Fixed Asset Turnover Ratio for All Companies')
plt.xlabel('Company Name')
plt.ylabel('Fixed Asset Turnover Ratio')

# Show the plot
plt.show()



# Set the figure size
plt.figure(figsize=(14, 8))

# Create a bar plot for Asset Turnover Ratio
sns.barplot(x='Company Name', y='Working Capital Turnover Ratio', data=filtered_df, palette='magma')

# Rotate the x labels for better readability
plt.xticks(rotation=90)

# Set the title and labels
plt.title('Working Capital Turnover Ratio for All Companies')
plt.xlabel('Company Name')
plt.ylabel('Working Capital Turnover Ratio')

# Show the plot
plt.show()