import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Generate sample data (replace with your own normalized data)
data = {
    'Category': ['A', 'B', 'C', 'D'],
    'DonationAmount': [500, 1000, 750, 900],
    'ConstituentCount': [50, 70, 60, 80]
}

# Create a Pandas DataFrame from the data
df = pd.DataFrame(data)

# Create a new workbook
workbook = Workbook()
writer = pd.ExcelWriter('output.xlsx', engine='openpyxl')
writer.book = workbook

# Add the DataFrame to the workbook on Sheet2
df.to_excel(writer, sheet_name='Sheet2', index=False)

# Create the different graphs
fig, axes = plt.subplots(nrows=2, ncols=2, figsize=(10, 8))

# Bar Chart - Donation Amount
axes[0, 0].bar(df['Category'], df['DonationAmount'])
axes[0, 0].set_title('Donation Amount')
axes[0, 0].set_xlabel('Category')
axes[0, 0].set_ylabel('Amount')

# Bar Chart - Constituent Count
axes[0, 1].bar(df['Category'], df['ConstituentCount'])
axes[0, 1].set_title('Constituent Count')
axes[0, 1].set_xlabel('Category')
axes[0, 1].set_ylabel('Count')

# Pie Chart - Donation Source Distribution
axes[1, 0].pie(df['DonationAmount'], labels=df['Category'], autopct='%1.1f%%')
axes[1, 0].set_title('Donation Source Distribution')

# Line Chart - Donation Trend
axes[1, 1].plot(df['Category'], df['DonationAmount'], marker='o')
axes[1, 1].set_title('Donation Trend')
axes[1, 1].set_xlabel('Category')
axes[1, 1].set_ylabel('Amount')

# Adjust the layout and spacing
plt.tight_layout()

# Save the figure to the workbook on Sheet2
worksheet = writer.sheets['Sheet2']
worksheet.add_image(plt.gcf(), 'D1')

# Save the workbook
writer.save()
