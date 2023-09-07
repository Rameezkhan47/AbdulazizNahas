import pandas as pd


def read_value_from_excel(filename, column="E", row=5):
    

    # Read the Excel file into a pandas DataFrame
    excel_file = pd.ExcelFile(filename)

    # Access a specific sheet by name
    sheet_name = 'Live Trading'  # Replace with the name of the sheet you want to access
    df = excel_file.parse(sheet_name)

    # Access a specific cell value
    row_index = 5  # Replace with the row index of the cell you want to access
    column_index = 5  # Replace with the column index of the cell you want to access
    cell_value = df.iloc[row_index-2, column_index-1]
    excel_file.close()
    return cell_value


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
        return stock_values
    else:
        print(f"Stock '{stock_name}' not found in the Excel file.")
        return []
    

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
            new_values = values[1:-1]

            df.iloc[:, stock_column_index] = new_values

            # Save the updated DataFrame back to the Excel file
            df.to_excel(filename, index=False)
            print(f"Values written to '{values[0]}' column in '{filename}'")
        else:
            print(f"Stock '{values[0]}' not found in the Excel file.")

    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == '__main__':
    filename = './resources/Chart Setting.xlsx'
    stock_name = "NVDA"
    values = [48, 23, 'Checkasdasd', 2, 'HAB asdasd Extreme', 48, 'Open', 'HABasasd High']
    write_stock_data_to_excel(filename, stock_name, values)


    
