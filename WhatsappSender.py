import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, Toplevel
import threading, gspread, os, sys, requests, json
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime


# Global variables for message parameters
message_template, term_in_tamil, due_date = "", "", ""
settings_file = "parameters.json"


def resource_path(relative_path):
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Global variable to control stopping of the process
stop_requested = threading.Event()

# Function to save parameters to a JSON file
def save_parameters_to_file():
    global message_template, term_in_tamil, due_date
    parameters = {
        "message_template": message_template,
        "term_in_tamil": term_in_tamil,
        "due_date": due_date
    }
    with open(resource_path(settings_file), 'w') as f:
        json.dump(parameters, f)

# Function to load parameters from the JSON file
def load_parameters_from_file():
    global message_template, term_in_tamil, due_date
    if os.path.exists(resource_path(settings_file)):
        with open(resource_path(settings_file), 'r') as f:
            parameters = json.load(f)
            message_template = parameters.get("message_template", message_template)
            term_in_tamil = parameters.get("term_in_tamil", term_in_tamil)
            due_date = parameters.get("due_date", due_date)
            
load_parameters_from_file()

# Function to authenticate and open the Google Sheets (return entire spreadsheet if sheet_name is None)
def open_google_sheet(sheet_url, sheet_name=None):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(resource_path("whatsapp.json"), scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_url(sheet_url)
    
    # Return the entire spreadsheet if no specific sheet is provided
    if sheet_name:
        sheet = spreadsheet.worksheet(sheet_name)
        return sheet
    return spreadsheet  # Return the full spreadsheet if no sheet name is given

def send_facebook_message(phone_number, due_fees, student_name):
    global message_template, term_in_tamil, due_date
    url = "https://graph.facebook.com/v20.0/159339593939407/messages"
    
    # Read Facebook Access Token from whatsapp.json
    with open(resource_path("whatsapp.json"), 'r') as f:
        config = json.load(f)
        ACCESS_TOKEN = config['facebook_access_token']

    # Conditional logic: Exclude term_in_tamil if message_template is "hostel_fees_due_tamil"
    parameters = [
        {"type": "text", "text": student_name},
        {"type": "text", "text": due_fees},
        {"type": "text", "text": due_date}
    ]
    
    # Add term_in_tamil only if template is not "hostel_fees_due_tamil"
    if (message_template != "hostel_fees_due_tamil" and message_template != "total_due_fess_tamil"):
        parameters.insert(1, {"type": "text", "text": term_in_tamil})

    payload = {
        "messaging_product": "whatsapp",
        "recipient_type": "individual",
        "to": phone_number,
        "type": "template",
        "template": {
            "name": message_template,
            "language": {
                "code": "ta"
            },
            "components": [
                {
                    "type": "body",
                    "parameters": parameters
                }
            ]
        }
    }
    
    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.post(url, headers=headers, json=payload)
        response.raise_for_status()  # Raise an error if the request failed
        return response.json()  # Returns the response as JSON
    except requests.exceptions.RequestException as e:
        return {"error": str(e)}

# Function to log messages into Google Sheets and GUI
def log_message(sheet, student, phone_number, grade, section, due_fees, message_id, response, treeview):
    current_datetime = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    sheet.append_row([
        current_datetime,
        student,
        phone_number,
        grade,
        section,
        due_fees,
        message_id,
        json.dumps(response)
    ])
    # Update the GUI Treeview with the message details
    new_item = treeview.insert('', 'end', values=[current_datetime, student, phone_number, due_fees, grade, section, message_id])
    treeview.see(new_item)
# Main function to extract student information and send messages
def extract_student_info(status_label, progress_var, progress_bar, treeview):
    # Set the initial state of the progress bar to 0%
    progress_bar.set(0)
    progress_var.set("0.00%")
    status_label.configure(text="Starting...")
    
    sheet_url = "https://docs.google.com/spreadsheets/d/1HAL8ApMQMZqhH-8vUfQkobG9Ub38IWZ_bW0UrIDgHDA/edit#gid=0"
    sending_sheet = open_google_sheet(sheet_url, "Sending")
    testing_sheet = open_google_sheet(sheet_url, "Test")
    sent_sheet = open_google_sheet(sheet_url, "Sent")
    
    if checkbox_var.get():
        values = testing_sheet.get_all_values()
    else:
        values = sending_sheet.get_all_values()
    

    total_rows = len(values) - 1  # Subtract 1 to account for the header row
    processed_rows = 0
    
    for i in range(1, len(values)):
        if stop_requested.is_set():  # Check if stop has been requested
            status_label.configure(text="Process stopped.")
            progress_var.set("0.00%")
            progress_bar.set(0)
            return
        
        row = values[i]
        student_name, phone_number1, phone_number2, phone_number3, grade, section, due_fees = row[1], row[2], row[3], row[4], row[5], row[6], row[7]

        # Collect phone numbers and remove duplicates
        phone_numbers = list(filter(None, [phone_number1.replace("-", ""), phone_number2.replace("-", ""), phone_number3.replace("-", "")]))
        unique_phone_numbers = list(set(phone_numbers))  # Remove duplicates
        
        for phone_number in unique_phone_numbers:
            response = send_facebook_message(phone_number, due_fees, student_name)
            message_id = response.get('messages', [{}])[0].get('id', 'No ID')
            log_message(sent_sheet, student_name, phone_number, grade, section, due_fees, message_id, response, treeview)
        
        processed_rows += 1
        progress_percentage = (processed_rows / total_rows) * 100
        progress_var.set(f"{progress_percentage:.2f}%")
        progress_bar.set(progress_percentage / 100)  # Update the progress bar
        status_label.configure(text=f"Processed row {i}/{total_rows}")
        
        # Update the UI to reflect changes
        status_label.update_idletasks()
        progress_bar.update_idletasks()

    status_label.configure(text="Process completed.")
    progress_var.set("100.00%")
    progress_bar.set(1)  # Set progress bar to 100% at the end
    stop_extraction()

# Function to start the extraction process in a new thread
def start_extraction(status_label, progress_var, progress_bar, treeview):
    global stop_requested
    stop_requested.clear()  # Clear the stop request
    start_button.configure(state=tk.DISABLED)
    edit_button.configure(state=tk.DISABLED)
    export_button.configure(state=tk.DISABLED)
    stop_button.configure(state=tk.NORMAL)
    threading.Thread(target=extract_student_info, args=(status_label, progress_var, progress_bar, treeview)).start()

# Function to stop the extraction process
def stop_extraction():
    global stop_requested
    # Disable the stop button after stopping the process
    stop_button.configure(state=tk.DISABLED)
    export_button.configure(state=tk.NORMAL)
    start_button.configure(state=tk.NORMAL)
    edit_button.configure(state=tk.NORMAL)
    stop_requested.set()  # Set the stop request
    
# Function to clear the Treeview
def clear_log(treeview):
    for item in treeview.get_children():
        treeview.delete(item)


def show_toast(root, message, duration=2000):
    # Create a new CTkToplevel window for the toast
    toast = ctk.CTkToplevel(root)
    toast.overrideredirect(True)  # Remove window decorations
    toast.attributes("-topmost", True)  # Keep it on top
    toast.attributes("-alpha", 0.95)  # Slight transparency

    # Calculate position to center the toast in the parent window
    window_width = 320
    window_height = 90
    x = root.winfo_x() + (root.winfo_width() // 2) - (window_width // 2)
    y = root.winfo_y() + (root.winfo_height() // 2) - (window_height // 2)
    toast.geometry(f"{window_width}x{window_height}+{x}+{y}")

    # Create a seamless frame with no borders and smooth rounded corners
    frame = ctk.CTkFrame(toast, corner_radius=10, fg_color=["#f2f2f2", "#333333"], border_width=0)
    frame.pack(fill='both', expand=True)

    # Create a label to display the message with transparent background and centered text
    toast_label = ctk.CTkLabel(frame, text=message, text_color=["#333333", "#e0e0e0"], 
                               fg_color="transparent", font=("Arial", 14), anchor="center")
    toast_label.pack(fill='both', expand=True, padx=20, pady=10)

    # Close the toast after the set duration (default: 2000ms)
    toast.after(duration, toast.destroy)

# Function to fetch templates and terms from a Google Sheet
def fetch_templates_and_terms():
    # Open the Google Sheets document
    sheet_url = "https://docs.google.com/spreadsheets/d/1HAL8ApMQMZqhH-8vUfQkobG9Ub38IWZ_bW0UrIDgHDA/edit#gid=0"  # Replace with your actual sheet URL
    spreadsheet = open_google_sheet(sheet_url)

    # Fetch templates (assuming they are in the first sheet 'Variables' and column A)
    template_sheet = spreadsheet.worksheet("Variables")
    templates = [row[0] for row in template_sheet.get_all_values()[1:] if row[0]]  # Skipping the header and excluding empty cells

    # Fetch terms (assuming they are in the second column of the same sheet)
    terms = [row[1] for row in template_sheet.get_all_values()[1:] if row[1]]  # Skipping the header and excluding empty cells


    return templates, terms

# Function to edit message parameters
def edit_parameters(template_label, term_label, date_label):
    global message_template, term_in_tamil, due_date

    # Fetch templates and terms from Google Sheets
    templates, terms = fetch_templates_and_terms()

    # Function to save parameters when "Save" button is clicked
    def save_parameters(option_menu_template, option_menu_term, entry_date):
        global message_template, term_in_tamil, due_date
        message_template = option_menu_template.get()  # Get the selected template from the dropdown
        term_in_tamil = option_menu_term.get()  # Get the selected term from the dropdown
        due_date = entry_date.get()

        # Save updated parameters to file
        save_parameters_to_file()

        # Update the labels with new values
        template_label.configure(text=message_template)
        term_label.configure(text=term_in_tamil)
        date_label.configure(text=due_date)

        # Show toast notification
        show_toast(root, "Parameters updated successfully!")

        # Close the parameters window
        params_window.destroy()

    # Create a new window for editing parameters
    params_window = ctk.CTkToplevel()
    params_window.title("Edit Message Parameters")
    params_window.geometry("500x350")
    params_window.attributes("-topmost", True)

    # Add the dropdown for message template selection
    ctk.CTkLabel(params_window, text="Message Template:", font=('Arial', 12)).pack(pady=5)
    option_menu_template = ctk.CTkOptionMenu(params_window, values=templates)  # Use templates from Google Sheet
    option_menu_template.pack(pady=5)
    option_menu_template.set(message_template)  # Set the current template as selected

    # Add the dropdown for term in Tamil
    ctk.CTkLabel(params_window, text="Term in Tamil", font=('Arial', 12)).pack(pady=5)
    option_menu_term = ctk.CTkOptionMenu(params_window, values=terms)  # Use terms from Google Sheet
    option_menu_term.pack(pady=5)
    option_menu_term.set(term_in_tamil)  # Set the current term as selected

    # Add the entry for the due date
    ctk.CTkLabel(params_window, text="Due Date:", font=('Arial', 12)).pack(pady=5)
    entry_date = ctk.CTkEntry(params_window, font=('Arial', 12))
    entry_date.pack(pady=5)
    entry_date.insert(0, due_date)

    # Button to save the parameters
    ctk.CTkButton(params_window, text="Save", command=lambda: save_parameters(option_menu_template, option_menu_term, entry_date), font=('Arial', 12)).pack(pady=10)
    ctk.CTkButton(params_window, text="Cancel", command=lambda: params_window.destroy(), font=('Arial', 12)).pack(pady=10)


# Add a new function to view the history
# Function to view the history
def view_whatsapp_history():
    
    def load_history():
        # Open the 'whatsapp message history' spreadsheet
        history_url = "https://docs.google.com/spreadsheets/d/1n6E9MCSOazNp6Pth0iCvLvRnOgaxCrYplmxF830Eo2s/edit?gid=0#gid=0"  # Replace with actual URL or spreadsheet ID
        
        try:
            history_spreadsheet = open_google_sheet(history_url)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open spreadsheet: {e}")
            return
        
        # Open a new Toplevel window to display the history
        history_window = Toplevel(root)
        history_window.title("WhatsApp Message History")
        history_window.geometry("1200x600")
        history_window.attributes("-topmost", True)
        
        # Create a frame for the statistics and history view
        frame = ctk.CTkFrame(history_window)
        frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create a smaller frame for the statistics on the left
        stats_frame = ctk.CTkFrame(frame, width=200)  # Adjust width here
        stats_frame.pack(side='left', fill='y', padx=(0, 10))
        
        stats_columns = ["Date", "Template Name", "Rows"]
        stats_treeview = ttk.Treeview(stats_frame, columns=stats_columns, show='headings', height=30)
        
        # Set different column widths for each column in the statistics table
        stats_treeview.heading("Date", text="Date")
        stats_treeview.column("Date", width=100)  # Adjust the width for 'Date'
        
        stats_treeview.heading("Template Name", text="Template Name")
        stats_treeview.column("Template Name", width=200)  # Adjust the width for 'Template Name'
        
        stats_treeview.heading("Rows", text="Rows")
        stats_treeview.column("Rows", width=50)  # Adjust the width for 'Rows'
        
        stats_treeview.pack(fill='y', expand=True)
        
        # Create a larger frame for displaying history log on the right
        history_frame = ctk.CTkFrame(frame)
        history_frame.pack(side='right', fill='both', expand=True)
        
        history_columns = ["DateTime", "Student", "Phone Number", "Due Fees", "Grade", "Section", "Message ID"]
        history_treeview = ttk.Treeview(history_frame, columns=history_columns, show='headings', height=30)
        for col in history_columns:
            history_treeview.heading(col, text=col)
            history_treeview.column(col, width=column_widths[col], anchor='center')  # Adjust width and center-align

        history_treeview.pack(side='left', fill='both', expand=True)
        
        # Add scrollbars for the history Treeview
        history_scroll_y = ttk.Scrollbar(history_frame, orient="vertical", command=history_treeview.yview)
        history_scroll_y.pack(side='right', fill='y')
        history_treeview.configure(yscrollcommand=history_scroll_y.set)
        
        history_scroll_x = ttk.Scrollbar(history_frame, orient="horizontal", command=history_treeview.xview)
        history_scroll_x.pack(side='bottom', fill='x')
        history_treeview.configure(xscrollcommand=history_scroll_x.set)
        
        def load_statistics():
            # Clear existing statistics
            for item in stats_treeview.get_children():
                stats_treeview.delete(item)
            
            # Populate the statistics Treeview
            for sheet in history_spreadsheet.worksheets():
                title = sheet.title
                row_count = len(sheet.get_all_values()) - 1
                date = title.split(" ")[0]  # Modify this logic based on actual sheet naming convention
                template_name = title.split(" ")[1]  # Modify this logic based on actual sheet naming convention
                stats_treeview.insert('', 'end', values=[date, template_name, row_count])
        
        load_statistics()
        
        def load_sheet_data(event):
            selected_item = stats_treeview.selection()
            if not selected_item:
                messagebox.showerror("Error", "Please select a sheet.")
                return
            
            selected_title = (stats_treeview.item(selected_item)['values'][0] + " " + stats_treeview.item(selected_item)['values'][1])  # Assuming template name is used as the sheet title
            selected_worksheet = history_spreadsheet.worksheet(selected_title)
            history_values = selected_worksheet.get_all_values()
            
            # Clear the Treeview before inserting new data
            for item in history_treeview.get_children():
                history_treeview.delete(item)
            
            # Insert new rows
            for row in history_values[1:]:
                history_treeview.insert('', 'end', values=row)
        
        stats_treeview.bind('<<TreeviewSelect>>', load_sheet_data)
        stats_treeview.selection_set(stats_treeview.get_children()[0])
        load_sheet_data(None)
    
    threading.Thread(target=load_history).start()

# Function to export log to a new Google Sheet
def export_log_to_history(treeview):
    global message_template
    
    # Open the 'whatsapp message history' spreadsheet
    history_url = "https://docs.google.com/spreadsheets/d/1n6E9MCSOazNp6Pth0iCvLvRnOgaxCrYplmxF830Eo2s/edit?gid=0#gid=0"  # Replace with actual URL or spreadsheet ID
    try:
        # Open the entire spreadsheet (no specific sheet name is needed here)
        history_spreadsheet = open_google_sheet(history_url)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open spreadsheet: {e}")
        return
    
    # Create a new sheet with the current date and template name
    current_date = datetime.now().strftime("%d-%m-%Y")
    sheet_name = f"{current_date} {message_template}"
    
    try:
        # Create a new worksheet in the spreadsheet
        new_sheet = history_spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="10")
    except gspread.exceptions.APIError as e:
        messagebox.showerror("Error", f"Failed to create sheet: {e}")
        return
    
    # Write the headers to the new sheet
    headers = ["DateTime", "Student", "Phone Number", "Due Fees", "Grade", "Section", "Message ID"]
    new_sheet.append_row(headers)
    
    # Collect all rows from the Treeview
    rows = []
    for row_id in treeview.get_children():
        row_data = treeview.item(row_id)['values']
        rows.append(row_data)
    
    # Write all rows to the new sheet in one batch
    try:
        new_sheet.append_rows(rows, value_input_option='RAW')
    except gspread.exceptions.APIError as e:
        messagebox.showerror("Error", f"Failed to write rows: {e}")
        return
    
    messagebox.showinfo("Success", "Log exported to sheet succesfully")


# global message_template, term_in_tamil, due_date

ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
ctk.set_default_color_theme(resource_path("red.json"))  # Themes: "blue" (default), "green", "dark-blue"

root = ctk.CTk()
root.title("WhatsApp Message Sender")
root.geometry("1300x650")  # Increase window size for new labels
root.iconbitmap(resource_path("auro.ico"))

# Configure grid rows and columns for centering
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=1)
root.rowconfigure(0, weight=0)
root.rowconfigure(1, weight=0)
root.rowconfigure(2, weight=0)
root.rowconfigure(3, weight=0)
root.rowconfigure(4, weight=0)
root.rowconfigure(5, weight=0)
root.rowconfigure(6, weight=1)
root.rowconfigure(7, weight=0)
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=1)

# Progress and status labels
ctk.CTkLabel(root, text="Progress:", font=('Arial', 12)).grid(row=0, column=0, padx=10, pady=5, sticky='e')

progress_var = ctk.StringVar()
ctk.CTkLabel(root, textvariable=progress_var, font=('Arial', 12)).grid(row=0, column=1, padx=10, pady=5, sticky='w')

progress_bar = ctk.CTkProgressBar(root, orientation='horizontal')
progress_bar.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky='ew')
progress_bar.set(0)

status_label = ctk.CTkLabel(root, text="", font=('Arial', 12))
status_label.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky='w')

# Parameters frame to display the current parameters
parameters_frame = ctk.CTkFrame(root)
parameters_frame.grid(row=3, column=0, columnspan=2, pady=10)

ctk.CTkLabel(parameters_frame, text="Message Template:", font=('Arial', 12)).grid(row=0, column=0, padx=10, pady=5, sticky='e')
template_label = ctk.CTkLabel(parameters_frame, text=message_template, font=('Arial', 12))
template_label.grid(row=0, column=1, padx=10, pady=5, sticky='w')

ctk.CTkLabel(parameters_frame, text="Term in Tamil:", font=('Arial', 12)).grid(row=1, column=0, padx=10, pady=5, sticky='e')
term_label = ctk.CTkLabel(parameters_frame, text=term_in_tamil, font=('Arial', 12))
term_label.grid(row=1, column=1, padx=10, pady=5, sticky='w')

ctk.CTkLabel(parameters_frame, text="Due Date:", font=('Arial', 12)).grid(row=2, column=0, padx=10, pady=5, sticky='e')
date_label = ctk.CTkLabel(parameters_frame, text=due_date, font=('Arial', 12))
date_label.grid(row=2, column=1, padx=10, pady=5, sticky='w')

# Edit button centered in the frame
edit_button = ctk.CTkButton(parameters_frame, text="Edit Parameters", command=lambda: edit_parameters(template_label, term_label, date_label), font=('Arial', 12))
edit_button.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky='ew')  # Center the button using columnspan

# Frame for buttons
button_frame = ctk.CTkFrame(root)
button_frame.grid(row=5, column=0, columnspan=2, pady=30)
button_frame.columnconfigure(0, weight=1)
button_frame.columnconfigure(1, weight=1)
button_frame.columnconfigure(2, weight=1)
button_frame.columnconfigure(3, weight=1)
button_frame.columnconfigure(4, weight=1)  # Adding extra column for the new button

checkbox_var = ctk.BooleanVar(value=False)  # Default to checked
checkbox = ctk.CTkCheckBox(button_frame, text="Enable Sample Mode", variable=checkbox_var, onvalue=True, offvalue=False)
checkbox.grid(row=0, column=0, padx=10, pady=5, sticky='w')

start_button = ctk.CTkButton(button_frame, text="Start Sending", command=lambda: start_extraction(status_label, progress_var, progress_bar, treeview), font=('Arial', 12))
start_button.grid(row=0, column=1, padx=10)

stop_button = ctk.CTkButton(button_frame, text="Stop Sending", command=lambda: stop_extraction(), font=('Arial', 12), state=tk.DISABLED)
stop_button.grid(row=0, column=2, padx=10)

export_button = ctk.CTkButton(button_frame, text="Export Log", command=lambda: export_log_to_history(treeview), font=('Arial', 12), state=tk.DISABLED)
export_button.grid(row=0, column=3, padx=10)

view_history_button = ctk.CTkButton(button_frame, text="View History", command=view_whatsapp_history, font=('Arial', 12))
view_history_button.grid(row=0, column=4, padx=10)

# Treeview for displaying log
columns = ["DateTime", "Student", "Phone Number", "Due Fees", "Grade", "Section", "Message ID"]

column_widths = {
    "DateTime": 150,
    "Student": 150,
    "Phone Number": 120,
    "Due Fees": 100,
    "Grade": 80,
    "Section": 80,
    "Message ID": 200
}

treeview = ttk.Treeview(root, columns=columns, show='headings', height=20)
for col in columns:
    treeview.heading(col, text=col)
    treeview.column(col, width=column_widths[col], anchor='center')  # Adjust width and center-align

treeview.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

# Clear Log button placed on top of the Treeview
clear_button = ctk.CTkButton(root, text="Clear Log", command=lambda: clear_log(treeview), font=('Arial', 12))

# Place button in a fixed position over the Treeview (adjust x, y as needed)
clear_button.place(relx=0.92, rely=0.95, anchor='center')


root.mainloop()