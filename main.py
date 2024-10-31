import os
import tkinter as tk
from tkinter import filedialog, messagebox
import py7zr
import pandas as pd
from tkinter import ttk
import json

# Function to extract a 7z archive
def extract_7z(archive_path, extract_path, progress_bar):
    try:
        with py7zr.SevenZipFile(archive_path, mode='r') as z:
            total_files = len(z.getnames())
            progress_value = 0
            for file in z.getnames():
                z.extract(targets=[file], path=extract_path)
                progress_value += 1
                progress_bar['value'] = (progress_value / total_files) * 100
                progress_bar.update()
        messagebox.showinfo("Success", "7zip extraction completed!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to extract 7z file: {e}")

# Function to save data as JSON
def save_as_json(data, save_path, file_name):
    file_path = os.path.join(save_path, f"{file_name}.json")
    try:
        with open(file_path, 'w') as json_file:
            json.dump(data, json_file, indent=4)
        messagebox.showinfo("Success", f"Data saved as JSON: {file_name}.json")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save JSON: {e}")

# Function to save data as Excel workbook
def save_as_workbook(data, save_path, file_name):
    file_path = os.path.join(save_path, f"{file_name}.xlsx")
    try:
        df = pd.DataFrame(data)
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", f"Data saved as Workbook: {file_name}.xlsx")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save Workbook: {e}")

# Function to process and save data
def process_data(archive_path, save_folder, save_workbook, progress_bar, status_label):
    status_label.config(text="Extracting 7zip...")
    extract_7z(archive_path, save_folder, progress_bar)

    # Example data (this should be replaced by your actual data processing logic)
    example_data = [
        {"message": "Test message 1", "date": "2024-01-01", "sender": "+123456789"},
        {"message": "Test message 2", "date": "2024-01-02", "sender": "+987654321"}
    ]

    # Save data as JSON
    status_label.config(text="Saving data as JSON...")
    save_as_json(example_data, save_folder, "example_data")

    # If user requested workbook copy
    if save_workbook:
        status_label.config(text="Saving data as Workbook...")
        save_as_workbook(example_data, save_folder, "example_data")

    progress_bar['value'] = 100  # Complete progress
    status_label.config(text="Processing complete.")

# Function to browse for a 7zip file
def browse_7z():
    file_path = filedialog.askopenfilename(filetypes=[("7zip files", "*.7z")])
    if file_path:
        entry_7z.delete(0, tk.END)
        entry_7z.insert(0, file_path)

# Function to select a folder to save output
def browse_save_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_save_folder.delete(0, tk.END)
        entry_save_folder.insert(0, folder_path)

# Function to start processing when "Start" button is pressed
def start_processing():
    archive_path = entry_7z.get()
    save_folder = entry_save_folder.get()
    save_workbook = workbook_var.get()

    if not os.path.exists(archive_path):
        messagebox.showerror("Error", "Please select a valid 7zip file.")
        return

    if not os.path.exists(save_folder):
        messagebox.showerror("Error", "Please select a valid save folder.")
        return

    process_data(archive_path, save_folder, save_workbook, progress_bar, status_label)

# Create the main GUI window
root = tk.Tk()
root.title("iPhone Data Extraction and Processing Tool")

# Set the window size
root.geometry("500x350")

# 7zip file selection
label_7z = tk.Label(root, text="Select 7zip file:")
label_7z.pack(pady=5)

entry_7z = tk.Entry(root, width=50)
entry_7z.pack(padx=10, pady=5)

btn_browse_7z = tk.Button(root, text="Browse", command=browse_7z)
btn_browse_7z.pack(pady=5)

# Save folder selection
label_save_folder = tk.Label(root, text="Select folder to save extracted data:")
label_save_folder.pack(pady=5)

entry_save_folder = tk.Entry(root, width=50)
entry_save_folder.pack(padx=10, pady=5)

btn_browse_save_folder = tk.Button(root, text="Browse", command=browse_save_folder)
btn_browse_save_folder.pack(pady=5)

# Option to save workbook copy
workbook_var = tk.BooleanVar()
workbook_checkbox = tk.Checkbutton(root, text="Save as Workbook", variable=workbook_var)
workbook_checkbox.pack(pady=5)

# Start button
btn_start = tk.Button(root, text="Start", command=start_processing)
btn_start.pack(pady=10)

# Progress bar
progress_bar = ttk.Progressbar(root, length=400, mode='determinate')
progress_bar.pack(pady=10)

# Status label
status_label = tk.Label(root, text="Waiting to start...", relief=tk.SUNKEN, anchor=tk.W)
status_label.pack(fill=tk.X, padx=10, pady=5)

# Run the Tkinter main loop
root.mainloop()
