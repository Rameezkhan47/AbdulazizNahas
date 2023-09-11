import glob
import os
from time import sleep
from os import listdir
from datetime import datetime
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import ElementClickInterceptedException, ElementNotInteractableException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
import xlwings as xw
import stock_data as data


def get(driver, css, wait=0, wait_for=True, attr='', return_text=False):
    try:
        elem = WebDriverWait(driver, wait).until(EC.presence_of_element_located((By.CSS_SELECTOR, css)))
        if return_text:
            return elem.text
        if attr:
            return elem.get_attribute(attr)
        else:
            return elem
    except:
        if wait_for:
            raise Exception(f'{css} not found')
        else:
            return ''


def get_all(driver, css):
    return driver.find_elements(By.CSS_SELECTOR, css)


def click(driver, css="", element="", element_wait=0, wait=2):
    try:
        element = get(driver, css, element_wait) if css else element
        element.click()

    except (ElementClickInterceptedException, ElementNotInteractableException):
        print(f'Element click intercepted, retrying')

        try:
            ActionChains(driver).move_to_element(element).click(element).perform()
            print('retry click successfull')

        except (ElementClickInterceptedException, ElementNotInteractableException):
            print(f'Element click intercepted again, retrying')
            driver.execute_script("arguments[0].click();", element)
            print('retry click successfull')

    sleep(wait)


def login(driver, username, password, retry_count=0):
    if retry_count == 3:
        print("failed to login.")
        driver.quit()
        exit()

    # removing popup if present
    while True:
        if get(webdriver, "[aria-label=Close]", 15, False):
            click(webdriver, ".close-icon-FuMQAaGA")
        if get(webdriver, "[aria-label=Close]", wait_for=False):
            webdriver.refresh()
        else:
            break

    # signing in to the website
    click(webdriver, ".menu-U2jIw4km")
    signin = [elem for elem in get_all(webdriver, '.label-jFqVJoPk') if "sign in" in elem.text.lower()][0]
    click(webdriver, element=signin)
    click(webdriver, '[name="Email"]', element_wait=5)
    username_div = get(webdriver, "#id_username", 5)
    username_div.send_keys(username)
    password_div = get(webdriver, "#id_password")
    password_div.send_keys(password)
    click(webdriver, ".submitButton-LQwxK8Bm")

    while get(driver, "#g-recaptcha-response", 5, False):
        if get(driver, ".antigate_solver.solved", 120, False):
            click(driver, ".submitButton-LQwxK8Bm", wait=4)
            continue
        else:
            webdriver.refresh()
            login(driver, username, password, ++retry_count)


def time_interval_parser(time_in_hours):
    print("time_in_hours: ", time_in_hours)
    print("time_in_hours: ", type(time_in_hours))
    # time_interval = int(time_in_hours)
    time_interval = time_in_hours
    is_in_days = time_in_hours // 24

    return time_interval*60, f"{is_in_days}D" if is_in_days else is_in_days


def extract_chart_data(webdriver, stock, time_interval, t3s_period, t3s_type,
                       PHPL_points, RSHVB_source, RSHVB_time_frame,
                       JFPCCI_source, t3v_source):
    # selecting stock
    counter = 0
    click(webdriver, ".wrap-IEe5qpW4")
    while not get(webdriver, f'[data-symbol-short="{stock}"]', wait_for=False):
        for i in range(10):
            actions.send_keys(Keys.DOWN)
        actions.perform()
        counter += 1
        if counter == 30:
            print(f"{stock} was not found")
            return True

    while get(webdriver, f'[data-symbol-short="{stock}"][data-active="false"]', wait_for=False):
        click(webdriver, f'[data-symbol-short="{stock}"]', wait=5)

    # selecting days interval
    click(webdriver, '#header-toolbar-intervals [aria-label="Time Interval"]', element_wait=10)
    time_interval, is_in_days = time_interval_parser(time_interval)
    if is_in_days:
        days_tab = [elem for elem in get_all(webdriver, ".accessible-raQdxQp0") if "days" in elem.text.lower()][0]
        click(webdriver, element=days_tab) if days_tab.get_attribute("data-open") == "false" else ""
        click(webdriver, f'[data-name="popup-menu-container"] [data-value="{is_in_days}"]')
    else:
        hours_tab = [elem for elem in get_all(webdriver, ".accessible-raQdxQp0") if "hours" in elem.text.lower()][0]
        click(webdriver, element=hours_tab) if hours_tab.get_attribute("data-open") == "false" else ""
        click(webdriver, f'[data-name="popup-menu-container"] [data-value="{time_interval}"]')

    # setting t3s settings
    actions.move_to_element(get_all(webdriver, '[data-name=legend-source-description]')[1]).perform()
    click(webdriver, '[data-name="legend-settings-action"]')
    click(webdriver, "#inputs")
    click(webdriver, '[data-section-name="T3S Settings"]')
    click(webdriver, ".cell-tBgV1m0B:has(#in_99) + .cell-tBgV1m0B + .cell-tBgV1m0B")
    actions.send_keys()
    actions.send_keys(Keys.BACKSPACE).perform()
    actions.send_keys(f"{t3s_period}").perform()

    click(webdriver, "#in_109")
    click(webdriver, "#in_102")
    click(webdriver, f"#id_in_102_item_T3-{t3s_type}", element_wait=5)
    click(webdriver, '[data-name="submit-button"]')

    # setting pivot high and low settings
    actions.move_to_element(get_all(webdriver, '[data-name=legend-source-description]')[2]).perform()
    click(webdriver, element=get_all(webdriver, '[data-name="legend-settings-action"]')[1])
    click(webdriver, "#inputs")
    click(webdriver, '[inputmode="numeric"]')
    actions.send_keys()
    actions.send_keys(Keys.BACKSPACE).perform()
    actions.send_keys(f"{PHPL_points}").perform()
    click(webdriver, element=get_all(webdriver, '[inputmode="numeric"]')[1])
    actions.send_keys()
    actions.send_keys(Keys.BACKSPACE).perform()
    actions.send_keys(f"{PHPL_points}").perform()
    click(webdriver, '[data-name="submit-button"]')

    # setting RSHVB settings
    actions.move_to_element(get_all(webdriver, '[data-name=legend-source-description]')[3]).perform()
    click(webdriver, element=get_all(webdriver, '[data-name="legend-settings-action"]')[2])
    click(webdriver, "#inputs")
    click(webdriver, "#in_1")
    click(webdriver, '[id="id_in_1_item_HAB-Trend-Biased-(Extreme)"]', element_wait=5)
    while get(webdriver, '#in_11[aria-expanded="false"]', wait_for=False):
        click(webdriver, "#in_11")
    time_interval, is_in_days = time_interval_parser(RSHVB_time_frame)
    click(webdriver, f'[id="id_in_11_item_{is_in_days if is_in_days else time_interval}"]')
    click(webdriver, '[data-name="submit-button"]')

    # setting JFPCCI settings
    actions.move_to_element(get_all(webdriver, '[data-name=legend-source-description]')[4]).perform()
    click(webdriver, element=get_all(webdriver, '[data-name="legend-settings-action"]')[3])
    click(webdriver, "#inputs")
    click(webdriver, "#in_1")
    click(webdriver, "#id_in_1_item_open", element_wait=5)
    click(webdriver, '[data-name="submit-button"]')

    # setting velocity settings
    actions.move_to_element(get_all(webdriver, '[data-name=legend-source-description]')[5]).perform()
    click(webdriver, element=get_all(webdriver, '[data-name="legend-settings-action"]')[4])
    click(webdriver, "#inputs")
    click(webdriver, "#in_1")
    click(webdriver, "#id_in_1_item_HAB-High", element_wait=5)
    click(webdriver, '[data-name="submit-button"]')

    # going back to 2018
    actions.click_and_hold(get(webdriver,  ".time-axis")).perform()
    for i in range(100):
        actions.move_by_offset(6, 0).perform()
    actions.release().perform()
    sleep(10)

    # downloading data
    click(webdriver, '[data-name="save-load-menu"]')
    click(webdriver, '[data-name="save-load-menu-item-clone"]+[data-role="menuitem"]')
    click(webdriver, '#time-format-select')
    click(webdriver, '#time-format-iso')
    click(webdriver, '[data-name="submit-button"]')

def extract_chart_data_with_retry(webdriver, stock, time_interval, t3s_period, t3s_type,
                                  PHPL_points, RSHVB_source, RSHVB_time_frame,
                                  JFPCCI_source, t3v_source):
    MAX_RETRIES = 6  # Set the maximum number of retries

    for _ in range(MAX_RETRIES):
        try:
            extract_chart_data(webdriver, stock, time_interval, t3s_period, t3s_type,
                               PHPL_points, RSHVB_source, RSHVB_time_frame,
                               JFPCCI_source, t3v_source)
            break  # If the function completes without errors, break out of the loop
        except Exception as e:
            print(f"An error occurred: {str(e)}")
            print("Reloading the page and retrying...")
            webdriver.refresh()  # Reload the page
            sleep(5)  # Wait for some time before retrying

def excel_functions(stock):
    # Clearing data in the excel file
    excel_file = f"./resources/{stock}.xlsm"
    
    wb = xw.Book(excel_file)
    ws = wb.sheets['Data Placement']
    ws.activate()
    app = wb.app
    # into brackets, the path of the macro
    data_clear_macro = app.macro(f"'{stock}.xlsm'!Module5.Delete_Data")
    live_trading_macro = app.macro(f"'{stock}.xlsm'!Module6.RunAna2")
    save_close_macro = app.macro(f"'{stock}.xlsm'!Module4.Save_Close")
    # data_clear_macro()
    print("done clearing the data")

    # Reading and filtering the downloaded data and copying to excel
    download_path = 'C:\\Users\\aamna\\Downloads'
    csv_file = [f"{download_path}/{p}" for p in listdir(download_path) if p.endswith(".csv") and stock in p]
    df = pd.read_csv(csv_file[0])
    if csv_file:
        df = pd.read_csv(csv_file[0])
    else:
        print("No matching CSV files found.")
        return
    stock_data = data.get_stock_data(stock)
    get_date = stock_data[-1]
    print(get_date)
    date = get_date.strftime('%Y-%m-%d')
    print("date is: ", date)
    filtered_df = df[df['time'] >= date]

    ws.range('C2').value = filtered_df.values.tolist()
    print("Filtered CSV saved successfully.")

    # live_trading_macro()

    # sleep(10)
    # save_close_macro()
    # sleep(10)
    wb.save()  # Save the workbook
    wb.close()  # Close the workbook
    app.quit()
    
    print("here")
    for filename in os.listdir(download_path):
            if stock.upper() in filename:
                file_path = os.path.join(download_path, filename)
                if os.path.isfile(file_path):
                    os.remove(file_path)
                    print(f"Deleted file: {file_path}")




# # Specify the path to the Excel file and the sheet name
# excel_file_path = './resources/Chart Setting.xlsx'
# sheet_name = 'Testing Log'  # Change this to the name of the sheet you want to read

# # Read the data from the specified sheet into a DataFrame
# tikker_data = pd.read_excel(excel_file_path)
# tikker_data = {str(tikker): tikker_data[tikker] for tikker in tikker_data.keys()}

# option = ChromeOptions()
# # option.add_argument("--headless=new")
# option.add_extension('./plugin.zip')
# option.add_argument('--no-sandbox')
# option.add_argument('--disable-dev-shm-usage')
# option.add_argument("--enable-logging")
# option.add_argument("--log-level=3")
# option.add_argument("--net-log-level=3")
# option.add_argument("--start-maximized")
# webdriver = Chrome(options=option)
# webdriver.get("https://www.tradingview.com/chart/CKrVzyF2/")
# actions = ActionChains(webdriver)

# # signing in to the website
# login(webdriver, "Azooz.nas@gmail.com", "Azooz0500306968!")
# webdriver.refresh()
# sleep(5)

# # moving draggable popup to bottom of the page so that it doesn't obstruct clicks
# if draggable_popup := get(webdriver, ".ui-draggable", wait_for=False):
#     js_move_element = """
#     var elementToMove = arguments[0];
#     var targetLocation = arguments[1];
#     targetLocation.appendChild(elementToMove);
#     """

#     webdriver.execute_script(js_move_element, draggable_popup,
#                              get(webdriver, '[aria-label="Notifications"]', 10))

# # opening watchlist if not opened and loading all stocks
# if get(webdriver, '[aria-label="Watchlist, details and news"][aria-pressed="false"]', wait_for=False):
#     click(webdriver, '[aria-label="Watchlist, details and news"]')

# # opening settings if hidden
# if hidden_settings := get(webdriver,
#                           ".closed-l31H9iuA .objectsTreeCanBeShown-l31H9iuA .iconArrow-l31H9iuA",
#                           wait_for=False):
#     click(webdriver, element=hidden_settings)


# def run_analysis(stock_name):
#     print(*tikker_data[stock_name][:-2])
#     extract_chart_data(webdriver, stock_name, *tikker_data[stock_name])
#     excel_functions(stock_name)

#     # webdriver.quit()


while (user_input := input("Enter Stock name to extract data or type \"exit\" to turn off program:\t")) != "exit":
    # error = extract_chart_data(webdriver, user_input, *tikker_data[str(user_input)][:-2])
    # # print(*tikker_data[str(user_input)][:-2])
    # if error:
    #     continue
    excel_functions(user_input)

# webdriver.quit()
