import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from math import pi

# Load the Excel file
file_path = 'Include/CEO.xlsx'
df = pd.read_excel(file_path, sheet_name='CEO Letter Analysis')

# Strip any leading/trailing spaces from column names
df.columns = [c.strip() for c in df.columns]

# Correct the column name if necessary
df.columns = df.columns.str.replace('SMOG  Index', 'SMOG Index')

# Display the first few rows of the DataFrame
print(df.head())

# Ensure the relevant columns exist in the DataFrame
integrated_thinking_topics = [
    'Board and CEO drive Integrated Thinking adoption', 
    'Integrated strategy', 
    'Culture of trust and innovation', 
    'Information system dicx'
]

# Check if all required columns exist in the DataFrame
missing_columns = [col for col in integrated_thinking_topics if col not in df.columns]
if missing_columns:
    raise ValueError(f"Missing columns in DataFrame: {missing_columns}")

# Number of variables
num_vars = len(integrated_thinking_topics)

# Compute angle for each axis
angles = [n / float(num_vars) * 2 * pi for n in range(num_vars)]
angles += angles[:1]  # Complete the loop

plt.figure(figsize=(10, 10))

ax = plt.subplot(111, polar=True)

for i, row in df.iterrows():
    # Drop the non-numeric column and select only the sustainability topics
    values = row[integrated_thinking_topics].values.flatten().tolist()
    values += values[:1]  # Complete the loop
    ax.plot(angles, values, linewidth=2, linestyle='solid', label=row['Company Name'])
    ax.fill(angles, values, alpha=0.25)

# Add labels and title
ax.set_xticks(angles[:-1])
ax.set_xticklabels(integrated_thinking_topics)
plt.title('Radar Chart of iIntegrated Thinking Topics in CEO Letters')
plt.legend(loc='lower right', bbox_to_anchor=(0.0, -0.15))
plt.show()
