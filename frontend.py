import sys
from PyQt6.QtWidgets import QApplication,QHeaderView,QComboBox, QMainWindow, QTableWidget, QTableWidgetItem, QPushButton, QVBoxLayout, QWidget, QDialog, QLabel, QHBoxLayout
from PyQt6.QtCore import (Qt, pyqtSignal)
from PyQt6.QtWidgets import QApplication,QLineEdit, QDialog, QVBoxLayout, QLabel, QTableWidget, QTableWidgetItem, QComboBox,QDateEdit
from PyQt6.QtCore import QTimer, QDateTime
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.action_chains import ActionChains
import trading_view
import stock_data
import json
import os
import threading
from queue import Queue
import time
from datetime import datetime



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
stock_queue = Queue()

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
        timeframes = stock_data.read_all_column_values(1)
        t3s_period = stock_data.read_all_column_values(2)
        t3s_type = stock_data.read_all_column_values(3)
        pivot = stock_data.read_all_column_values(4)
        rshvb_source = stock_data.read_all_column_values(5)
        rshvb_time = stock_data.read_all_column_values(6)
        jfpcci = stock_data.read_all_column_values(7)
        t3v = stock_data.read_all_column_values(8)

        # Populate the settings table with rows
        settings_data = [

            ("Stock Name", [stock]),
            ("Chart Time Frame Type", [str(data[0]),*timeframes]),
            ("Algorithmic Decipher - T3S Period", [str(data[1]),*t3s_period]),
            ("Algorithmic Decipher - T3S Type", [str(data[2]),*t3s_type]),
            ("Pivot High Low Points - Left and Right Bars",[str(data[3]),*pivot]),
            ("RSHVB - Source", [str(data[4]),*rshvb_source] ),
            ("RSHVB - Time Frame    ", [str(data[5]),*rshvb_time] ),
            ("JFPCCI - Source", [str(data[6]),*jfpcci]),
            ("T3V - Source", [str(data[7]),*t3v]),
            ("Timer to next analysis", [str(data[8]),"5 seconds", "10 seconds", "3 day", "2 day", "1 day", "12 hours", "6 hours", "4 hours", "3 hours", "2 hours", "1 hour", "3 seconds"]),
            ("Start date filter on the extracted data", [QDateEdit(data[9])]),

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
        date_format = "%d-%b-%Y"
        values[-1] = datetime.strptime(values[-1], date_format)
        stock_data.write_stock_data_to_excel(values)





class CountdownWidget(QWidget):
    def __init__(self, stock, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        self.label = QLabel("1:00:00", self)
        self.stock_name = stock
        layout.addWidget(self.label)
        time = stock_data.get_stock_data(stock)
        time = time[8]
        # time = '2 seconds'
        
        
        if isinstance(time, int):
            time = ((time)*3600)  
            # time = time
        elif 'hour' in time or 'hours' in time:
            time = (int(time.split()[0])*3600)
            
        elif 'minute' in time or 'minutes' in time:
            time = (int(time.split()[0]) * 60)
        elif 'second' in time or 'seconds' in time:
            time = (int(time.split()[0]))
        elif 'day' in time or 'days' in time:
            time = (int(time.split()[0]) * (24*60*60))
        self.time = time
        self.remaining_seconds = int(time)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_countdown)
        self.timer.start(1000)  # Update every second
        self.update_countdown()

    def update_countdown(self):
        if (self.remaining_seconds) > 0:
            hours = self.remaining_seconds // 3600
            minutes = (self.remaining_seconds % 3600) // 60
            seconds = self.remaining_seconds % 60
            time_str = "{:02d}:{:02d}:{:02d}".format(hours, minutes, seconds)
            self.label.setText(time_str)
            self.remaining_seconds -= 1
        else:
            self.timer.stop()
            self.enqueue_for_analysis()
            self.reset_timer()  # Call the method to reset the timer
         

    def enqueue_for_analysis(self):
        stock_queue.put(self.stock_name)
        
    def reset_timer(self):
        # Reset the timer back to the initial time (2 seconds in this case)
        # self.remaining_seconds = self.time
        self.remaining_seconds = self.time
        
        self.timer.start(1000)  # Start the timer again
        
class AnalysisHistoryPopup(QDialog):
    def __init__(self, stock, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Analysis History for {}".format(stock))
        layout = QVBoxLayout(self)
        
        # Load historical data for the specified stock from history.json
        history_data = self.load_history_data(stock)
        
        # Display the historical data
        self.history_label = QLabel("Historical analysis results for {}:".format(stock))
        layout.addWidget(self.history_label)
        
        # Create a table or any layout for historical analysis data here
        # For this example, let's just add a label
        self.history_data_label = QLabel(self.format_history_data(history_data))
        layout.addWidget(self.history_data_label)
    
    def load_history_data(self, stock):
        try:
            with open(f"./resources/history.json", "r") as json_file:
                data = json.load(json_file)
                return data.get(stock, {})
        except FileNotFoundError:
            return {}
    
    def format_history_data(self, data):
        formatted_data = ""
        for timestamp, value in data.items():
            formatted_data += "{}: {}\n".format(timestamp, value)
        return formatted_data


class AnalysisResultWidget(QWidget):
    def __init__(self, stock, parent=None):
        super().__init__(parent)
        self.stock = stock  # Store the stock symbol
        layout = QHBoxLayout(self)
        self.analysis_result_label = QLabel(self.get_latest_analysis_result())
        self.history_button = QPushButton("History")
        self.history_button.setMinimumHeight(30)
        self.history_button.clicked.connect(self.open_analysis_history_popup)
        layout.addWidget(self.analysis_result_label)
        layout.addWidget(self.history_button)
        self.stock = stock

    def get_latest_analysis_result(self):
        # print(self.stock)
        # return
        return stock_data.read_value_from_excel(f'./resources/{self.stock}.xlsm', column=7, row=4)

    def update_analysis_result_label(self):
        print("Inside update_analysis_result_label")
        latest_result = self.get_latest_analysis_result()
        print("latest_result: ", latest_result)
        self.analysis_result_label.setText(latest_result)

    def open_analysis_history_popup(self):
        history_popup = AnalysisHistoryPopup(self.stock)
        history_popup.exec()
        
class StockNameWidget(QWidget):
    def __init__(self, analysis_result_widget, parent=None, stock_name=""):
        super().__init__(parent)
        self.analysis_result_widget = analysis_result_widget  # Store the reference to the AnalysisResultWidget
        self.stock_name = stock_name  # Store the stock_name as an instance variable
        layout = QHBoxLayout(self)
        self.analysis_result_label = QLabel(stock_name)
        layout.addWidget(self.analysis_result_label)



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
        return
        # Retrieve the updated last_price value from the Excel file
        last_price = stock_data.read_value_from_excel(self.stock_name)
        
        # Convert the integer to a string and update the last_price label with the new value
        self.last_price_label.setText(str('100'))



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
        self.analysis_result_widgets = {}
        self.load_data()
        
    def process_stock_queue(self):
        while True:
            stock_name = stock_queue.get()
            print(stock_queue)
            self.run_analysis(stock_name)
    def run(self):
        # Start a separate thread to process the stock queue
        thread = threading.Thread(target=self.process_stock_queue)
        thread.daemon = True  # This thread will exit when the main program exits
        thread.start()


    def load_data(self):
        # Set the number of rows in the table
        self.table.setRowCount(len(stocks))  # Example: 10 rows

        # Populate the table with data from your sources
        for row, stock in enumerate(stocks):
            last_price_item = LastPriceWidget(excel_file=f"./resources/TSLA.xlsm")
            change_percent_item = QTableWidgetItem("+2.5%")

            # Create the custom widget for the analysis result
            analysis_result_widget = AnalysisResultWidget(stock)
            self.analysis_result_widget = analysis_result_widget
            stock_name_item = StockNameWidget(analysis_result_widget, stock_name=stock)
            
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
            countdown_widget = CountdownWidget(stock, parent=self)
            self.table.setCellWidget(row, 4, countdown_widget)
            self.table.setCellWidget(row, 5, edit_settings_button)
            self.table.setRowHeight(row, 50)
            self.analysis_result_widgets[stock] = analysis_result_widget


    def run_analysis(self, stock_name):
        print("Running analysis for: ", stock_name)
        trading_view.run_analysis(stock_name)
        analysis_result_widget = self.analysis_result_widgets.get(stock_name)

        if analysis_result_widget:
            analysis_result_widget.update_analysis_result_label()
        time.sleep(3)  # Pause for 5 seconds
        return True
        # Call the test function with the stock_name
       
    def open_chart_settings_popup(self, stock):
        chart_settings_popup = ChartSettingsPopup(stock, parent=self)
        chart_settings_popup.exec()

def main():
    app = QApplication(sys.argv)
    window = StockApp()
    window.show()
    window.run()  # Start the queue processing thread
    sys.exit(app.exec())

if __name__ == "__main__":
    main()