import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = 'Include/ratios.xlsx'
dataframes = pd.read_excel(file_path, sheet_name=None)


# Extract the specific sheet
df = dataframes['Sheet1']

# Display the first few rows of the dataframe to understand its structure
print(df.head())

columns_of_interest = ['Company Name', 'Net working capital', 'Net working capital.1', 'Net working capital.2', 'Net working capital.3', 'Net working capital.4', 'Net working capital.5', 'Net working capital.6', 'Net working capital.7', 'Net working capital.8', 'Net working capital.9', 'Net working capital.10']
df_negative_wc = df[columns_of_interest]

df_negative_wc.columns = ['Company Name', '2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023']

# Convert data to numeric, coerce errors to NaN
df_negative_wc.iloc[:, 1:] = df_negative_wc.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')

# Display the cleaned dataframe
print(df_negative_wc.head())

output_file_path = 'cleaned_data.xlsx'
writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')
df_negative_wc.to_excel(writer, sheet_name='Negative Working Capital', index=False)
writer.close()
print(f"Cleaned data written to '{output_file_path}'")


#Liquidiy Ratios

# Calculate liquidity ratios for the selected companies
selected_companies = ['Chalet Hotels Ltd.', 'Lemon Tree Hotels Ltd.', 'T G B Banquets & Hotels Ltd.', 'Viceroy Hotels Ltd.', 'Taj G V K Hotels & Resorts Ltd.']

# Filter the dataframe for the selected companies
df_selected = df[df['Company Name'].isin(selected_companies)]

# Calculate the ratios
ratios = pd.DataFrame()
ratios['Company Name'] = df_selected['Company Name']

# Assuming Total Current Assets = Total Assets - Total Liabilities
ratios['Current Ratio'] = df_selected['Total assets'] / df_selected['Total liabilities']

# Assuming Quick Assets = Total Current Assets - Inventory (approximated by Raw material cycle)
ratios['Quick Ratio'] = (df_selected['Total assets'] - df_selected['Raw material cycle (days)']) / df_selected['Total liabilities']

# Cash Ratio
ratios['Cash Ratio'] = df_selected['Cash & bank balance to current assets'] / df_selected['Total liabilities']

print(ratios.head())

# Plot the ratios
plt.figure(figsize=(14, 8))

# Current Ratio
plt.subplot(3, 1, 1)
sns.barplot(x='Company Name', y='Current Ratio', data=ratios)
plt.title('Current Ratio')
plt.xticks(rotation=45)

# Quick Ratio
plt.subplot(3, 1, 2)
sns.barplot(x='Company Name', y='Quick Ratio', data=ratios)
plt.title('Quick Ratio')
plt.xticks(rotation=45)

# Cash Ratio
plt.subplot(3, 1, 3)
sns.barplot(x='Company Name', y='Cash Ratio', data=ratios)
plt.title('Cash Ratio')
plt.xticks(rotation=45)

plt.tight_layout()
plt.show()