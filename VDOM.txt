"""
FortiGate Configuration Generator Script

Author: Yogev Kahana
Version: Beta 1.0
Date: January 23, 2024

This script allows you to generate FortiGate configuration commands based on data from an Excel file.
The generated commands are then saved to a Notepad file.

Usage:
1. Run the script.
2. Select an Excel file containing the configuration data.
3. Choose the location and provide a name for the Notepad file to save the generated commands.

Note: This is a beta version of the script.

"""

import openpyxl
from tkinter import Tk, filedialog


def open_file_dialog():
    root = Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )

    return file_path


def save_file_dialog():
    root = Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.asksaveasfilename(
        title="Save the File",
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt")]
    )

    return file_path


def process_excel_file(excel_file_path, notepad_file_path):
    wb = openpyxl.load_workbook(excel_file_path)
    sheet = wb.active

    notepad_content = ""

    # Iterating through rows
    for row in sheet.iter_rows(min_row=2, values_only=True):
        vdomLink, name , type   = row

        # Building configuration commands
        config_commands = f"""
edit {vdomLink}
    set vdom {name}
    set type  {type }


next
"""

        # Appending the configuration commands to notepad_content
        notepad_content += config_commands

    # Adding "end" at the end of notepad_content
    notepad_content += "end"

    # Writing to Notepad
    with open(notepad_file_path, 'w') as notepad_file:
        notepad_file.write(notepad_content)

    print(f"Configuration commands written to {notepad_file_path}")


if __name__ == "__main__":
    excel_file_path = open_file_dialog()

    if excel_file_path:
        notepad_file_path = save_file_dialog()

        if notepad_file_path:
            process_excel_file(excel_file_path, notepad_file_path)
        else:
            print("No location selected to save the file.")
    else:
        print("No file selected.")
