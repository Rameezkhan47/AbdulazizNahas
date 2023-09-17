import datetime
import numpy as np
import pandas as pd


def read_value_from_excel(filename, column=5, row=5):
    try:
        print(filename)
        # Read the Excel file into a pandas DataFrame
        excel_file = pd.ExcelFile(filename)

        # Access a specific sheet by name
        sheet_name = 'Live Trading'  # Replace with the name of the sheet you want to access
        df = excel_file.parse(sheet_name)

        cell_value = df.iloc[row-2, column-1]
        excel_file.close()
        return str(cell_value)
    except Exception as e:
        return "No Value"  # Return the error message as a string
    
def read_all_column_values(column):
    try:
        # Read the Excel file into a pandas DataFrame
        excel_file = pd.ExcelFile('./resources/Dropdown for Chart Settings.xlsx')

        # Access a specific sheet by name
        sheet_name = 'Sheet1'  # Replace with the name of the sheet you want to access
        df = excel_file.parse(sheet_name)

        # Read all values from the specified column
        column_values = df.iloc[:, column-1].astype(str).tolist()
        column_values = [value for value in column_values if value != 'nan']
        column_values = [value.split('.')[0] for value in column_values]
        
        excel_file.close()
        return column_values
    except Exception as e:
        return []  # Return an empty list if there's an error
    
    
def get_stock_data(stock_name):
    # Load the Excel file into a DataFrame
    filename = './resources/Chart Setting.xlsx'
    df = pd.read_excel(filename)

    # Find the column index where the stock name is located
    stock_column_index = None
    for column in df.columns:
        if str(column) == str(stock_name):
            stock_column_index = df.columns.get_loc(column)
            break

    if stock_column_index is not None:
        # Extract all values under the stock name
        stock_values = df.iloc[:, stock_column_index].tolist()
        # print(stock_values)
        return stock_values
    else:
        print(f"Stock '{stock_name}' not found in the Excel file. Adding default data...")
        
        # Add a new column with default data
        default_data = ['48', '3', 'New', '2', 'HAB Trend Extreme', '48', 'Open', 'HAB High', 4, datetime.datetime(2018, 1, 4, 0, 0)]
        df[str(stock_name)] = default_data
        
        # Save the updated DataFrame back to the Excel file
        df.to_excel(filename, index=False)
        
        return default_data
    
def write_stock_data_to_excel(values):
    try:
        # Load the Excel file into a DataFrame
        filename = './resources/Chart Setting.xlsx'
        df = pd.read_excel(filename)

        # Find the column index where the stock name is located
        stock_column_index = None
        for column in df.columns:
            if str(column) == str(values[0]):
                stock_column_index = df.columns.get_loc(column)
                break

        if stock_column_index is not None:
            # Write the values to the specified column
            new_values = values[1:]

            df.iloc[:, stock_column_index] = new_values

            # Save the updated DataFrame back to the Excel file
            df.to_excel(filename, index=False)
            print(f"Values written to '{values[0]}' column in '{filename}'")
        else:
            print(f"Stock '{values[0]}' not found in the Excel file.")

    except Exception as e:
        print(f"An error occurrasdadasdasded: {str(e)}")
        
def get_formatted_value(input_string):
    modified_string = input_string.replace(' ', '-')
    return modified_string

    RSHVB_source = get_formatted_value(RSHVB_source)
    print("formatted RSHVB_source:  ",RSHVB_source)

    click(webdriver, f'[id="id_in_1_item_{RSHVB_source}"]', element_wait=5)


if __name__ == '__main__':
    all_values = read_all_column_values(1)
    print("result: ", all_values)


    
