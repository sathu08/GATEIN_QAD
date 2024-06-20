import time
import openpyxl
import pyautogui
from datetime import datetime
from pynput.keyboard import Controller, Key
from PySide2 import QtWidgets
import keyboard  # Import the keyboard module for key event handling
import logging
import os
import pandas as pd


# creating a instatance of keyboard
keyboard1 = Controller()
#  creating the log file
log_file = 'gatepass.log'
if not os.path.exists(log_file):
    with open(log_file, 'w'):
        pass

# Configure the logging settings
logging.basicConfig(filename=log_file, level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')


def search_excel(file_path, sheet_name, search_value, values_to_return):

    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]
        matching_values = []
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value == search_value:
                    matching_values = [row[ord(column) - ord('A')].value for column in values_to_return]
                    break
        workbook.close()
        print(matching_values)
        return matching_values
        logging.info("ASN search result"+str(matching_values))
    except Exception as e:
        logging.error(f"Error in search_excel: {str(e)}")
        return None

def wait_for_ctrl_j():    
    # Wait for the user to press Ctrl+J
    keyboard.wait('ctrl+j')

def handle_domestic_case(search_value, result_values, error_text):
    try:
        
        minimize_window()       
        wait_for_ctrl_j()  # Wait for Ctrl+J key press
        time.sleep(0.4)
        pyautogui.typewrite(result_values[0], interval=0.2)
        keyboard1.press(Key.tab)
        keyboard1.release(Key.tab)
        pyautogui.typewrite(str(result_values[1].upper()), interval=0.2)
        keyboard1.press(Key.tab)
        keyboard1.release(Key.tab)
        pyautogui.typewrite(result_values[2], interval=0.2)
        keyboard1.press(Key.enter)
        keyboard1.release(Key.enter)

        for _ in range(7):
            keyboard1.press(Key.tab)
            keyboard1.release(Key.tab)

        x = datetime.strptime(str(result_values[3]), '%d/%m/%Y')
        invdate = x.strftime('%m/%d/%Y')
        pyautogui.typewrite(invdate, interval=0.2)
        
        keyboard1.press(Key.enter)
        keyboard1.release(Key.enter)
        
        keyboard1.press(Key.enter)
        keyboard1.release(Key.enter)

    except Exception as e:
        show_error(f"Error in handle_domestic_case: {str(e)}",
                   (error_text + "BOE details not found in file"))


def handle_import_case(search_value, result_values,result_values_dos, error_text):
    try:
        
        minimize_window()
        wait_for_ctrl_j()  # Wait for Ctrl+J key press
        time.sleep(0.4)
        pyautogui.typewrite(result_values[0], interval=0.2)
        keyboard1.press(Key.tab)
        keyboard1.release(Key.tab)
        pyautogui.typewrite(str(result_values[1]), interval=0.2)
        keyboard1.press(Key.tab)
        keyboard1.release(Key.tab)
        pyautogui.typewrite(result_values[2], interval=0.2)
        keyboard1.press(Key.enter)
        keyboard1.release(Key.enter)
        for _ in range(7):
            keyboard1.press(Key.tab)
            keyboard1.release(Key.tab)
        sub_date = datetime.strptime(str(result_values[3]), '%d/%m/%Y')
        invdate = sub_date.strftime('%m/%d/%Y')
        pyautogui.typewrite(invdate, interval=0.1)       
        if result_values_dos:
            keyboard1.press(Key.tab)
            keyboard1.release(Key.tab)
            pyautogui.typewrite(str(result_values_dos[0]), interval=0.1)
            keyboard1.press(Key.tab)
            keyboard1.release(Key.tab)
            date_object = datetime.strptime(str(result_values_dos[1]), "%Y-%m-%d %H:%M:%S")
            time_string = date_object.strftime("%m/%d/%Y")
            pyautogui.typewrite(time_string, interval=0.2)
            keyboard1.press(Key.tab)
            keyboard1.release(Key.tab)
            pyautogui.typewrite(result_values_dos[2], interval=0.2)
            keyboard1.press(Key.enter)
            keyboard1.release(Key.enter)
            time.sleep(1)
        else:
            error_message = "BOE details not found"
            show_error(error_message, error_text)        
    except Exception as e:
        show_error(f"Error in handle_not_domestic_case: {str(e)}", error_text)


def show_error(message, error_text):
    try:
        error_text.insertPlainText(message + '\n')  # Append the message to the text widget
        logging.error(message)
    except Exception as e:
        logging.error(f"Error updating error_text: {str(e)}")


def clear_error(error_text):
    error_text.clear()


def maximize_window():
    pass


def minimize_window():
    window.showMinimized()


def get_search_value():
    search_value = entry.text()
    file_path1 = "ASN.xlsx"
    sheet_name1 = "Sheet1"
    values_to_return = ["E", "A", "D", "P"]
    if search_value:
        result_values1 = search_excel(file_path1, sheet_name1, search_value, values_to_return)
        if radio_domestic.isChecked():
            if result_values1:
                handle_domestic_case(search_value, result_values1, error_text)
            else:
                error_message = "Value not found."
                show_error(error_message, error_text)
                logging.error(f"Value not found for search_value1: {search_value}")
        elif radio_import.isChecked():
            file_path_dos = "workfile.xlsx"
            sheet_name_dos = "VIC"
            search_value_dos = str(search_value)
            values_to_return_dos = ["A", "B", "T"]
            result_values_dos = search_excel(file_path_dos, sheet_name_dos, search_value_dos, values_to_return_dos)
            # Handle import case here
            handle_import_case(search_value, result_values1,result_values_dos, error_text)
            print(result_values_dos)
        else:
            error_message = "Please select Domestic or Import."
            show_error(error_message, error_text)
    else:
        # Handle the case where no value is entered
        error_message = "Please enter a search value."
        show_error(error_message, error_text)
        logging.warning("No search value entered by the user")


def file_copy():
    try:
        with open("path.txt") as f:
            path = f.read()
            # file_cust_fil = path
    except Exception as e:
        error_message = f"Error: {e}"
        show_error(error_message, error_text)
    input_file = path
    df = pd.read_excel(input_file, skiprows=2)
    output_file = "workfile.xlsx"
    # Convert 'BOE Date' column to datetime
    df['BOE Date'] = pd.to_datetime(df['BOE Date'], errors='coerce')
    today = datetime.today()
    three_months_ago = today - pd.DateOffset(months=3)  # Change months to 3 for three months ago
    filtered_df = df[df['BOE Date'] >= three_months_ago]
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        filtered_df.to_excel(writer, sheet_name='VIC', index=False)
    print("Filtered data saved to", output_file)
    error_message = f"Filtered data saved to: {output_file}"
    show_error(error_message, error_text)

          
app = QtWidgets.QApplication([])

# GUI window for user input
window = QtWidgets.QWidget()
window.setWindowTitle("GATE PASS")
layout = QtWidgets.QVBoxLayout()
# label for invoice field
label = QtWidgets.QLabel("Enter Invoice Number:")
layout.addWidget(label)
# textbox for invoice field
entry = QtWidgets.QLineEdit()
layout.addWidget(entry)
# radio buttons for Domestic and Import cases
radio_domestic = QtWidgets.QRadioButton("Domestic")
radio_import = QtWidgets.QRadioButton("Import")
# button group for radio buttons
radio_button_group = QtWidgets.QButtonGroup()
radio_button_group.addButton(radio_domestic)
radio_button_group.addButton(radio_import)
#  radio buttons to the layout
layout.addWidget(radio_domestic)
layout.addWidget(radio_import)
# submit pushbutton
submit_button = QtWidgets.QPushButton("Submit")
submit_button.clicked.connect(get_search_value)
layout.addWidget(submit_button)
# error  box
error_text = QtWidgets.QTextEdit()
layout.addWidget(error_text)
# clear button
clear_button = QtWidgets.QPushButton("Clear Errors")
clear_button.clicked.connect(lambda: clear_error(error_text))
layout.addWidget(clear_button)
# load import details
load_button = QtWidgets.QPushButton("Load Import Details")
load_button.clicked.connect(lambda: file_copy())
layout.addWidget(load_button)
window.setLayout(layout)
window.show()
app.exec_()
