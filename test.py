import xlwings as xw
import pandas as pd

excel_file = "./resources/TSLA.xlsm"

wb = xw.Book(excel_file)
ws = wb.sheets['Data Placement']
app = wb.app
# into brackets, the path of the macro
macro_vba = app.macro("'TSLA.xlsm'!Module5.Delete_Data")
macro_vba()


print("done")

# Read the CSV file into a DataFrame
csv_file = '/Users/apple/Downloads/BATS_MSFT, 2D.csv'
df = pd.read_csv(csv_file)

# Convert the "time" column to datetime format
df['time'] = pd.to_datetime(df['time'])

# Filter rows from 2018 onwards
filtered_df = df[df['time'] >= '2018-01-01']


ws.range('C2').value = filtered_df.values.tolist()

print("Filtered CSV saved successfully.")

wb.save()  # Save the workbook
wb.close()  # Close the workbook
app.quit()
