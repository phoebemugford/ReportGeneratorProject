import tkinter as tk
from tkinter import filedialog, ttk, messagebox, font
from tkinter import *
import time
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util import Retry
import traceback

from spacesOverview import spaces_overview_report
from deviceReport import device_overview_report
from monthlyTrends import monthly_trends_report

AUTH = "Authentication Key"
URL = "API.com"
HEADERS = {"Authorization": f"Bearer {AUTH}"}
SESSION = requests.Session()
RETRY = Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
SESSION.mount('http://', HTTPAdapter(max_retries=RETRY))
SESSION.mount('https://', HTTPAdapter(max_retries=RETRY))

# Create the GUI window
window = tk.Tk()
window.title("Report Generator")
window.geometry("400x350")  # Set the window size


def open_directory_dialog():
    root = tk.Tk()
    root.withdraw()
    directory_path = filedialog.askdirectory()
    save_location_entry.delete(0, tk.END)
    save_location_entry.insert(0, directory_path)  # Display the selected directory


# Runs code for selected report(s)
def generate_reports(spaces_overview, device_readings, monthly_trends, save_location, spaces_to_include, work_hours,
                     start_date, end_date, year_chosen):
    try:
        spaces_df = get_spaces(spaces_to_include)
        adjusted_window_height = window.winfo_height() + 55  # Adjusts window size for progress bar
        window.geometry(f'400x{adjusted_window_height}')
        number_of_reports = spaces_overview + device_readings + monthly_trends
        report_on = 0
        if spaces_overview == 1:
            report_on = report_on + 1
            report_track_string = "Generating Report " + str(report_on) + "/" + str(number_of_reports)
            spaces_overview_report(save_location, spaces_to_include, spaces_df, window, report_track_string)
        if device_readings == 1:
            report_on = report_on + 1
            report_track_string = "Generating Report " + str(report_on) + "/" + str(number_of_reports)
            device_overview_report(save_location, start_date, end_date, work_hours, spaces_to_include, spaces_df, window,
                                   report_track_string)
        if monthly_trends == 1:
            report_on = report_on + 1
            report_track_string = "Generating Report " + str(report_on) + "/" + str(number_of_reports)
            monthly_trends_report(save_location, year_chosen, work_hours, spaces_to_include, spaces_df, window, report_track_string)
        adjusted_window_height = window.winfo_height() - 55  # Readjusts window size
        window.geometry(f'400x{adjusted_window_height}')
        messagebox.showinfo("Report Generation Complete", "Report generation complete!")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        traceback.print_exc()


# Function to get spaces to be tested
def get_spaces(spaces_to_include):
    url = f"{URL}spaces"
    response = SESSION.get(url, headers=HEADERS)
    print(response.status_code)
    data_json = response.json()
    df = pd.DataFrame(data_json["spaces"], columns=["id", "name"])
    df = df.dropna().reset_index(drop=True)
    spaces_to_ignore = []
    with open('IgnoredSpaces') as f:  # Retrieves spaces in IgnoredSpaces file
        for line in f:
            id_to_append = line.split(" ", 1)[0]
            spaces_to_ignore.append(id_to_append)
    if len(spaces_to_include) > 0:
        # Adds all spaces without user-specified text to list of ignored spaces
        spaces = data_json["spaces"]
        for space in spaces:
            id_to_append = space.get('id')
            name_to_append = space.get('name')
            if not id_to_append in spaces_to_ignore and not spaces_to_include in name_to_append:
                spaces_to_ignore.append(id_to_append)
    df = df[~df['id'].isin(spaces_to_ignore)]  # Removes all spaces in spaces_to_ignore from testing list
    # spaces_to_test = ['spc_pb5fz8w0', 'spc_Bm8lYD26', 'spc_9dMswFLY']
    # df = df[df['id'].isin(spaces_to_test)]
    return df


# Disables generate button until save location and at least one report is selected
def update_generate_button_state(*args):
    if save_location_var.get() and (spaces_overview_var.get() == 1 or device_readings_var.get() == 1 or
                                    monthly_trends_var.get() == 1):
        generate_button.config(state=tk.NORMAL)
    else:
        generate_button.config(state=tk.DISABLED)


# Function to enable/disable the space selection box
def update_space_selection_state(*args):
    if all_spaces_var.get() == 0:
        space_entry.config(state=tk.NORMAL)
    else:
        space_entry.config(state=tk.DISABLED)
        start_date_label.grid_remove()


# Shows additional parameters for device report when report is selected, hides otherwise
def show_device_report_options(*args):
    if device_readings_var.get() == 1:
        start_date_label.grid()
        start_date_entry.grid()
        end_date_label.grid()
        end_date_entry.grid()
        adjusted_window_height = window.winfo_height() + 100  # Resizes window
        window.geometry(f'400x{adjusted_window_height}')
    else:
        start_date_label.grid_remove()
        start_date_entry.grid_remove()
        end_date_label.grid_remove()
        end_date_entry.grid_remove()
        adjusted_window_height = window.winfo_height() - 100  # Resizes window
        window.geometry(f'400x{adjusted_window_height}')


# Shows additional parameters for monthly trends report when report is selected, hides otherwise
def show_monthly_trends_options(*args):
    if monthly_trends_var.get() == 1:
        current_year_check.grid()
        year_entry_label.grid()
        year_entry.grid()
        adjusted_window_height = window.winfo_height() + 84  # Resizes window
        window.geometry(f'400x{adjusted_window_height}')
    else:
        current_year_check.grid_remove()
        year_entry_label.grid_remove()
        year_entry.grid_remove()
        adjusted_window_height = window.winfo_height() - 84  # Resizes window
        window.geometry(f'400x{adjusted_window_height}')


# Disables year selection entry until current year box is unchecked
def update_year_selection_state(*args):
    if current_year_var.get() == 0:
        year_entry.config(state=tk.NORMAL)
    else:
        year_entry.config(state=tk.DISABLED)


# Creates variables for widgets
save_location_var = tk.StringVar()
spaces_overview_var = tk.IntVar(value=0)
device_readings_var = tk.IntVar(value=0)
monthly_trends_var = tk.IntVar(value=0)
start_date_var = tk.StringVar()
end_date_var = tk.StringVar()
work_hours_only_var = tk.IntVar()
all_spaces_var = tk.IntVar(value=1)  # Starts checked
space_entry_var = tk.StringVar()
current_year_var = tk.IntVar(value=1)  # Starts checked
year_entry_var = tk.StringVar()

# Creates widgets
select_location_button = tk.Button(window, text="Select Location", command=open_directory_dialog)
select_location_button.grid(column=0, row=0, pady=3)

save_location_label = tk.Label(window, text="Save Location:")
save_location_label.grid(column=0, row=1, pady=3)

save_location_entry = tk.Entry(window, width=40, textvariable=save_location_var)
save_location_entry.grid(column=0, row=2, padx=75, pady=3)

all_spaces_check = tk.Checkbutton(window, text='Include All Spaces', variable=all_spaces_var, onvalue=1,
                                  offvalue=0)
all_spaces_check.grid(column=0, row=3, pady=3)

space_entry_label = tk.Label(window, text="Only Include Spaces with Following Text:")
space_entry_label.grid(column=0, row=4)

all_spaces_var.trace('w', update_space_selection_state)
space_entry = tk.Entry(window, width=40, textvariable=space_entry_var, state=tk.DISABLED)
space_entry.grid(column=0, row=5, pady=3)

work_hours_only_check = tk.Checkbutton(window, text='Include Work Hours Only', variable=work_hours_only_var, onvalue=1,
                                       offvalue=0)
work_hours_only_check.grid(column=0, row=6, pady=3)

report_options_label = tk.Label(window, text="Select Reports to Generate:")
report_options_label.grid(column=0, row=7, pady=3)

spaces_overview_check = tk.Checkbutton(window, text='Current Overview of Spaces', variable=spaces_overview_var, onvalue=1,
                                       offvalue=0)
spaces_overview_check.grid(column=0, row=8, pady=3)

device_readings_check = tk.Checkbutton(window, text='Device Readings Overview', variable=device_readings_var, onvalue=1,
                                       offvalue=0)
device_readings_check.grid(column=0, row=9, pady=3)
device_readings_var.trace('w', show_device_report_options)  # Binds device report check to additional parameters

start_date_label = tk.Label(window, text="Start Date for Device Readings (dd/mm/yyyy):")
start_date_label.grid(column=0, row=10, pady=3)
start_date_label.grid_remove()  # Hides widget until device report is selected

start_date_entry = tk.Entry(window, width=40, textvariable=start_date_var)
start_date_entry.grid(column=0, row=11, pady=3)
start_date_entry.grid_remove()  # Hides widget until device report is selected

end_date_label = tk.Label(window, text="End Date for Device Readings (dd/mm/yyyy):")
end_date_label.grid(column=0, row=12, pady=3)
end_date_label.grid_remove()  # Hides widget until device report is selected

end_date_entry = tk.Entry(window, width=40, textvariable=end_date_var)
end_date_entry.grid(column=0, row=13, pady=3)
end_date_entry.grid_remove()  # Hides widget until device report is selected

monthly_trends_check = tk.Checkbutton(window, text='Monthly Trends', variable=monthly_trends_var, onvalue=1,
                                      offvalue=0)
monthly_trends_check.grid(column=0, row=14, pady=3)
monthly_trends_var.trace('w', show_monthly_trends_options)  # Binds monthly trends report check to additional parameters

current_year_check = tk.Checkbutton(window, text='Generate Monthly Trends for Current Year', variable=current_year_var, onvalue=1,
                                    offvalue=0)
current_year_check.grid(column=0, row=15, pady=3)
current_year_check.grid_remove()  # Hides widget until monthly trends report is selected

year_entry_label = tk.Label(window, text="Year to Analyze for Monthly Trends (yyyy):")
year_entry_label.grid(column=0, row=16, pady=3)
year_entry_label.grid_remove()  # Hides widget until monthly trends report is selected

current_year_var.trace('w', update_year_selection_state)
year_entry = tk.Entry(window, width=40, textvariable=year_entry_var, state=tk.DISABLED)
year_entry.grid(column=0, row=17, pady=3)
year_entry.grid_remove()  # Hides widget until monthly trends report is selected

# Binds the variables to the generate button
save_location_var.trace('w', update_generate_button_state)
spaces_overview_var.trace('w', update_generate_button_state)
device_readings_var.trace('w', update_generate_button_state)
monthly_trends_var.trace('w', update_generate_button_state)

generate_button = tk.Button(window, text="Generate Reports",
                            command=lambda: generate_reports(spaces_overview_var.get(), device_readings_var.get(),
                                                             monthly_trends_var.get(), save_location_entry.get(),
                                                             space_entry.get(), work_hours_only_var.get(),
                                                             start_date_entry.get(), end_date_entry.get(),
                                                             year_entry_var.get()),
                            state=tk.DISABLED)
generate_button.grid(column=0, row=18, pady=3)

# Start the GUI event loop
window.mainloop()
