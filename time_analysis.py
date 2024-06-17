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

# Let's extract the columns for the years 2013 to 2023 and rename them accordingly
columns_of_interest = ['Company Name', 'Net working capital', 'Net working capital.1', 'Net working capital.2', 'Net working capital.3', 'Net working capital.4', 'Net working capital.5', 'Net working capital.6', 'Net working capital.7', 'Net working capital.8', 'Net working capital.9', 'Net working capital.10']
df_negative_wc = df[columns_of_interest]

# Rename columns for better readability
df_negative_wc.columns = ['Company Name', '2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023']

# Convert data to numeric, coerce errors to NaN
df_negative_wc.iloc[:, 1:] = df_negative_wc.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')

# Display the cleaned dataframe
print(df_negative_wc.head())


#Negative Working Capital

# Filter the data for the specific companies
companies = df[columns_of_interest[0]];
company_data = df_negative_wc[df_negative_wc['Company Name'].isin(companies)]

# Plot the trend for each company
plt.figure(figsize=(14, 8))
for company in companies:
    sns.lineplot(data=company_data[company_data['Company Name'] == company].iloc[0, 1:].T, marker='o', label=company)

plt.title('Negative Working Capital Trend for Selected Companies')
plt.xlabel('Year')
plt.ylabel('Net Working Capital')
plt.xticks(rotation=45)
# plt.legend()
plt.grid(True)
plt.show()