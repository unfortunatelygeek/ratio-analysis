import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = 'Include/CEO.xlsx'
df = pd.read_excel(file_path, sheet_name='CEO Letter Analysis')

# Strip any leading/trailing spaces from column names
df.columns = [c.strip() for c in df.columns]

# It seems there is an extra space in the column name 'SMOG Index'. Let's correct it.
df.columns = df.columns.str.replace('SMOG  Index', 'SMOG Index')

# Display the first few rows of the DataFrame
print(df.head())

# # 1. Distribution of Positive Sentiment across different industries
# plt.figure(figsize=(12, 6))
# sns.boxplot(x='Industry', y='Positive sentiment', data=df)
# plt.xticks(rotation=90)
# plt.title('Distribution of Positive Sentiment across Different Industries')
# plt.xlabel('Industry')
# plt.ylabel('Positive Sentiment')
# plt.show()

# # Ensure the relevant columns exist in the DataFrame
# readability_indices = [
#     'Gunning fog Index', 'SMOG Index', 'Flesch–Kincaid Index', 
#     'Coleman–Liau index', 'Automated readability index'
# ]

# # 1. Box Plot: Compare readability indices across different companies
# plt.figure(figsize=(15, 10))
# for i, index in enumerate(readability_indices, 1):
#     plt.subplot(2, 3, i)
#     sns.boxplot(x='Company Name', y=index, data=df)
#     plt.xticks(rotation=90)
#     plt.title(f'Comparison of {index} across Companies')
# plt.tight_layout()
# plt.show()

# # 2. Heatmap: Show the readability indices of different companies
# # Calculate the average readability indices for each company
# company_readability = df.groupby('Company Name')[readability_indices].mean().reset_index()

# plt.figure(figsize=(15, 10))
# sns.heatmap(company_readability.set_index('Company Name'), annot=True, cmap='coolwarm', center=0)
# plt.title('Heatmap of Average Readability Indices by Company')
# plt.show()

# # 3. Bar Plot: Compare the average readability indices of different companies
# company_avg_readability = df.groupby('Company Name')[readability_indices].mean().reset_index()

# plt.figure(figsize=(15, 10))
# for i, index in enumerate(readability_indices, 1):
#     plt.subplot(2, 3, i)
#     sns.barplot(x='Company Name', y=index, data=company_avg_readability)
#     plt.xticks(rotation=90)
#     plt.title(f'Average {index} by Company')
# plt.tight_layout()
# plt.show()

from math import pi

categories = ['Compliance', 'Business centred', 'Systematic', 'Regenerative', 'Coevolutionary']
num_vars = len(categories)

angles = [n / float(num_vars) * 2 * pi for n in range(num_vars)]
angles += angles[:1]

plt.figure(figsize=(8, 8))

ax = plt.subplot(111, polar=True)

for i, row in df.iterrows():
    values = row.drop('Company Name').values.flatten().tolist()
    values += values[:1]
    ax.plot(angles, values, linewidth=2, linestyle='solid', label=row['Company Name'])
    ax.fill(angles, values, alpha=0.25)

ax.set_xticks(angles[:-1])
ax.set_xticklabels(categories)
plt.title('Radar Chart of Sustainability Topics in CEO Letters')
plt.legend(loc='upper right', bbox_to_anchor=(0.1, 0.1))
plt.show()