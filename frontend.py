import sys
from PyQt6.QtWidgets import QApplication,QHeaderView,QComboBox, QMainWindow, QTableWidget, QTableWidgetItem, QPushButton, QVBoxLayout, QWidget, QDialog, QLabel, QHBoxLayout
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication,QLineEdit, QDialog, QVBoxLayout, QLabel, QTableWidget, QTableWidgetItem, QComboBox,QDateEdit
from PyQt6.QtCore import QTimer, QDateTime
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.action_chains import ActionChains
# import trading_view
import stock_data
import json
import os

def get_stocks():

    # Define the folder path where your Excel files are located
    folder_path = 'resources'

    # Initialize an empty list to store the stock names
    stock_names = []

    # List all files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsm') and not filename.startswith('~$'):
            # Extract the stock name from the file name (remove the '.xlsm' extension)
            stock_name = os.path.splitext(filename)[0]
            stock_names.append(stock_name)

    # Print the final list of stock names
    return stock_names

#get stocks from resources folder
stocks = get_stocks()

class AnalysisPopup(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Analysis Details")
        layout = QVBoxLayout(self)
        self.analysis_label = QLabel("Analysis details will be displayed here.")
        layout.addWidget(self.analysis_label)

class ChartSettingsPopup(QDialog):
    def __init__(self, stock, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Chart Settings")
        self.resize(800, 600)
        # print("stock: ",stock)
        layout = QVBoxLayout(self)
        data = stock_data.get_stock_data(stock)

        # Add a label for instructions
        self.settings_label = QLabel("Edit chart settings:")
        layout.addWidget(self.settings_label)

        # Create a table for chart settings
        self.settings_table = QTableWidget()
        self.settings_table.setColumnCount(2)
        self.settings_table.setHorizontalHeaderLabels(["Chart Indicator", "Selection/Input"])
        layout.addWidget(self.settings_table)

        # Populate the settings table with rows
        settings_data = [

            ("Stock Name", [stock]),
            ("Chart Time Frame Type", [str(data[0]),"3 day", "2 day", "1 day", "12 hours", "6 hours", "4 hours", "3 hours", "2 hours", "1 hour"]),
            ("Algorithmic Decipher - T3S Period", [str(data[1]),"3", "More Values..."]),
            ("Algorithmic Decipher - T3S Type", [str(data[2]),"New", "More Values..."]),
            ("Pivot High Low Points - Left and Right Bars",[str(data[3]),"2", "More Values..."]),
            ("RSHVB - Source", [str(data[4]),"HAB Trend Extreme", "More Values..."] ),
            ("RSHVB - Time Frame    ", [str(data[5]),"HAB Trend Extreme", "More Values..."] ),
            ("JFPCCI - Source", [str(data[6]),"Open", "More Values..."]),
            ("T3V - Source", [str(data[7]),"HAB High", "More Values"]),
            ("Start date filter on the extracted data", ["1-Jun-2018", "More Values..."]),
        ]

        self.settings_table.setRowCount(len(settings_data))
        for row, (indicator, choices) in enumerate(settings_data):
            indicator_item = QTableWidgetItem(indicator)
            self.settings_table.setItem(row, 0, indicator_item)
            
            if indicator == "Stock Name":
                # For "Stock Name" cell, use a non-editable QLineEdit
                text_input = QLineEdit()
                text_input.setText(choices[0])
                text_input.setReadOnly(True)  # Make it non-editable
                self.settings_table.setCellWidget(row, 1, text_input)
            elif isinstance(choices[0], QDateEdit):
                input_widget = choices[0]
                self.settings_table.setCellWidget(row, 1, input_widget)
            else:
                selection_item = QComboBox()
                selection_item.addItems(choices)
                self.settings_table.setCellWidget(row, 1, selection_item)

        table_width = self.settings_table.width()
        column_0_width = int(table_width * 0.4)  # 40% width for the first column
        column_1_width = table_width - column_0_width
        self.settings_table.setColumnWidth(0, column_0_width)
        self.settings_table.setColumnWidth(1, column_1_width)

        # Add the submit button
        submit_button = QPushButton("Submit")
        submit_button.clicked.connect(self.collect_and_print_data)
        layout.addWidget(submit_button)

    def collect_and_print_data(self):
        data = []
        for row in range(self.settings_table.rowCount()):
            indicator = self.settings_table.item(row, 0).text()
            widget = self.settings_table.cellWidget(row, 1)

            if isinstance(widget, QComboBox):
                value = widget.currentText()
            elif isinstance(widget, QLineEdit):
                value = widget.text()
            elif isinstance(widget, QDateEdit):
                value = widget.date().toString("dd-MMM-yyyy")
            else:
                value = ""

            data.append((indicator, value))
        values = [value for _, value in data]

        stock_data.write_stock_data_to_excel(values)





class CountdownWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        self.label = QLabel("1:00:00", self)
        layout.addWidget(self.label)
        
        self.remaining_seconds = 3600
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_countdown)
        self.timer.start(1000)  # Update every second
        self.update_countdown()

    def update_countdown(self):
        if self.remaining_seconds > 0:
            hours = self.remaining_seconds // 3600
            minutes = (self.remaining_seconds % 3600) // 60
            seconds = self.remaining_seconds % 60
            time_str = "{:02d}:{:02d}:{:02d}".format(hours, minutes, seconds)
            self.label.setText(time_str)
            self.remaining_seconds -= 1
        else:
            self.timer.stop()

class AnalysisHistoryPopup(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Analysis History")
        layout = QVBoxLayout(self)
        self.history_label = QLabel("Dummy historical analysis results:")
        layout.addWidget(self.history_label)
        # Create a table or any layout for historical analysis data here
        # For this example, let's just add a label
        self.history_data_label = QLabel("Historical data goes here.")
        layout.addWidget(self.history_data_label)

class AnalysisResultWidget(QWidget):
    def __init__(self, parent=None, analysis_result=""):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        self.analysis_result_label = QLabel(analysis_result)
        self.history_button = QPushButton("History")
        self.history_button.setMinimumHeight(30)
        self.history_button.clicked.connect(self.open_analysis_history_popup)
        layout.addWidget(self.analysis_result_label)
        layout.addWidget(self.history_button)

    def open_analysis_history_popup(self):
        history_popup = AnalysisHistoryPopup(self)
        history_popup.exec()
        
class StockNameWidget(QWidget):
    def __init__(self, parent=None, stock_name=""):
        super().__init__(parent)
        self.stock_name = stock_name  # Store the stock_name as an instance variable
        layout = QHBoxLayout(self)
        self.analysis_result_label = QLabel(stock_name)
        self.run_button = QPushButton("Run")
        self.run_button.setMinimumHeight(30)
        layout.addWidget(self.analysis_result_label)
        layout.addWidget(self.run_button)

        # Connect the "Run" button's clicked signal to the slot that runs the test function
        self.run_button.clicked.connect(lambda: self.run_test(stock_name.lstrip()))

    def run_test(self, stock_name):
        print("stock_name: ", stock_name)
        # trading_view.run_analysis(stock_name)
        
class LastPriceWidget(QWidget):
    def __init__(self, parent=None, excel_file=""):
        super().__init__(parent)
        self.stock_name = excel_file  # Store the stock_name as an instance variable
        layout = QHBoxLayout(self)
        self.last_price_label = QLabel("")
        self.refresh_button = QPushButton("Refresh")
        self.refresh_button.setMinimumHeight(30)
        layout.addWidget(self.last_price_label)
        layout.addWidget(self.refresh_button)

        # Connect the "Refresh" button's clicked signal to the slot that updates the last_price value
        self.refresh_button.clicked.connect(self.update_last_price)

        # Initialize the last_price label with the initial value
        self.update_last_price()

    def update_last_price(self):
        # Retrieve the updated last_price value from the Excel file
        last_price = stock_data.read_value_from_excel(self.stock_name)
        
        # Convert the integer to a string and update the last_price label with the new value
        self.last_price_label.setText(str(last_price))



class StockApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Stock Analysis App")
        self.resize(1280, 720)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        layout = QVBoxLayout(self.central_widget)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["Stock Name", "Last Stock Price", "Change %", "Latest Analysis Result", "Timer to Next Analysis", "Edit Chart Settings"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)  # Equal width for all columns
        self.table.verticalHeader().setVisible(False)  # Hide the row numbers
        self.table.setShowGrid(True)  # Show grid lines
        layout.addWidget(self.table)
        self.load_data()

    def load_data(self):
        # Set the number of rows in the table
        self.table.setRowCount(10)  # Example: 10 rows

        # Populate the table with data from your sources
        for row, stock in enumerate(stocks):
            stock_name_item = StockNameWidget(stock_name=stock)
            last_price_item = LastPriceWidget(excel_file=f"./resources/TSLA.xlsm")
            change_percent_item = QTableWidgetItem("+2.5%")

            # Create the custom widget for the analysis result
            analysis_result_widget = AnalysisResultWidget(analysis_result="Result {}".format(stock))
            edit_settings_button = QPushButton("Edit")
            edit_settings_button.clicked.connect(lambda _, stock=stock: self.open_chart_settings_popup(stock))


            # Set the items to non-editable
            # stock_name_item.setFlags(Qt.ItemFlag.NoItemFlags)
            # last_price_item.setFlags(Qt.ItemFlag.NoItemFlags)
            change_percent_item.setFlags(Qt.ItemFlag.NoItemFlags)

            self.table.setCellWidget(row, 0, stock_name_item)
            self.table.setCellWidget(row, 1, last_price_item)
            self.table.setItem(row, 2, change_percent_item)
            self.table.setCellWidget(row, 3, analysis_result_widget)
            countdown_widget = CountdownWidget()
            self.table.setCellWidget(row, 4, countdown_widget)
            self.table.setCellWidget(row, 5, edit_settings_button)
            self.table.setRowHeight(row, 50)

    def run_analysis(self, stock_item):
        stock_name = stock_item.text()
        # Call the test function with the stock_name
       
    def open_chart_settings_popup(self, stock):
        chart_settings_popup = ChartSettingsPopup(stock, parent=self)
        chart_settings_popup.exec()

def main():
    app = QApplication(sys.argv)
    window = StockApp()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
