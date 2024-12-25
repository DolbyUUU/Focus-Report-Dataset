import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# Read data from the Excel file
df = pd.read_excel('focus_report_index.xlsx')

# Convert the 'DATE' column to datetime
df['DATE'] = pd.to_datetime(df['DATE'])

# Set the date column as the index of the DataFrame
df.set_index('DATE', inplace=True)

# Calculate the monthly average of the 'INDEX' column
df_monthly = df.resample('M').mean()

# Create a time series plot of the monthly averages
plt.figure(figsize=(10,6))
plt.plot(df_monthly.index, df_monthly['INDEX'], marker='o', linestyle='-')
plt.xlabel('Year')
plt.ylabel('Monthly Average Index')
plt.title('Monthly Average Index Over Years')
plt.grid(True)

# Format x-ticks to show year
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
plt.gca().xaxis.set_major_locator(mdates.YearLocator()) 

plt.gcf().autofmt_xdate()  # autoformat the x-ticks for better display

# Save the plot as an image file
plt.savefig('monthly_average_index.png')

# Calculate the annual average of the 'INDEX' column
df_annual = df.resample('Y').mean()

# Create a time series plot of the annual averages
plt.figure(figsize=(10,6))
plt.plot(df_annual.index, df_annual['INDEX'], marker='o', linestyle='-')
plt.xlabel('Year')
plt.ylabel('Annual Average Index')
plt.title('Annual Average Index Over Years')
plt.grid(True)

# Format x-ticks to show year
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
plt.gca().xaxis.set_major_locator(mdates.YearLocator()) 

plt.gcf().autofmt_xdate()  # autoformat the x-ticks for better display

# Save the plot as an image file
plt.savefig('annual_average_index.png')