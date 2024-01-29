import openpyxl
from tkinter import Tk, filedialog, Button, Label
import tkinter as tk

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

def process_excel_file(excel_file_path, notepad_file_path, status_label_var):
    try:
        wb = openpyxl.load_workbook(excel_file_path)
    except Exception as e:
        status_label_var.set(f"Error: {e}")
        return

    notepad_content = ""

    # Iterate through the sheets in the workbook
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]

        # Building configuration commands for the current sheet
        sheet_commands = f"config system interface\n"  # Add the desired sentence at the beginning of each sheet
     #   if sheet_name == "vdom":  #Can add command for what sheet i will needed
      #      sheet_commands += "config global\n"

        # Check if the sentence was already added for this sheet
        sentence_added = False

        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not sentence_added:
                sentence_added = True

            if sheet_name == "interface":  # Replace with the actual name of the first sheet
                interface_name, vdom, mode, ip, security_mode, access = row

                # Building configuration commands for each row
                config_commands = f"""
edit {interface_name}
    set vdom {vdom}
    set mode {mode}
    set ip {ip}
    set security_mode {security_mode}
    set access {access}
next
"""

            elif sheet_name == "policy":  # Replace with the actual name of the sheet
                policy_id, name, srcaddr, srcintf, dstintf, dstaddr, service, action = row

                # Building configuration commands for each row
                config_commands = f"""
edit {policy_id}
    set name {name}
    set srcintf {srcintf}
    set dstintf {dstintf}
    set srcaddr {srcaddr}
    set dstaddr {dstaddr}
    set service {service}
    set action {action}
next
"""
            elif sheet_name == "object":  # Replace with the actual name of the  sheet
                name, subnet, comments = row

                # Building configuration commands for each row
                config_commands = f"""
edit {name}
    set subnet {subnet}
    set comments {comments}
next
"""
            elif sheet_name == "vlan": # Replace with the actual name of the sheet
                vlan, vdom, interface, type, vlanid, mode, ip, allowaccess = row

                # Building configuration commands
                config_commands = f"""
edit {vlan}
    set vdom {vdom}
    set interface  {interface}
    set type  {type}
    set vlanid  {vlanid}
    set mode  {mode}
    set ip  {ip}
    set allowaccess  {allowaccess}
next
"""
            elif sheet_name == "vdom":  # Replace with the actual name of the sheet
                vdomLink, vdomname, type = row

                # Building configuration commands
                config_commands = f"""
                config global\n
                edit {vdomLink}
                    set vdom {vdomname}
                    set type  {type}
next
"""
            elif sheet_name == "route":  # Replace with the actual name of the sheet
                number, destination, gateway, distance = row

                # Building configuration commands
                config_commands = f"""
                config router static \n
                edit {number}
                    set destination {destination}
                    set geteway {gateway}
                    set distance {distance}


next
"""
            sheet_commands += config_commands

        # Adding "end" at the end of notepad_content for each sheet
        sheet_commands += "end\n"

        # Adding the sheet commands to notepad_content
        notepad_content += sheet_commands

    # Write the notepad_content to a text file
    try:
        with open(notepad_file_path, 'w') as notepad_file:
            notepad_file.write(notepad_content)
        status_label_var.set(f"Configuration commands written to {notepad_file_path}")
    except Exception as e:
        status_label_var.set(f"Error: {e}")

def create_fortigate_rules():
    excel_file_path = open_file_dialog()

    if excel_file_path:
        notepad_file_path = save_file_dialog()

        if notepad_file_path:
            status_label_var = tk.StringVar()
            process_excel_file(excel_file_path, notepad_file_path, status_label_var)
            status_label.config(textvariable=status_label_var)
        else:
            status_label.config(text="No location selected to save the file.")
    else:
        status_label.config(text="No file selected.")

if __name__ == "__main__":
    window = tk.Tk()
    window.title("FortiGate Rule Generator")
    window.geometry("500x400")
    select_button = Button(window, text="Select Excel File", command=create_fortigate_rules)
    select_button.pack(pady=20)
    bottom_label = Label(window, text="@By Bynet Data Communications@", anchor='w', justify='left')
    bottom_label.pack(side='bottom', fill='x')
    status_label = Label(window, text="", anchor='w', justify='left')
    status_label.pack(side='bottom', fill='x')
    window.mainloop()
