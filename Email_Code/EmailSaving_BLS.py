import win32com.client as win32
import datetime
import tkinter as tk
from tkinter import messagebox
import os
from PyPDF2 import PdfMerger


def merge_pdfs_in_folder(folder_path, output_file):
    merger = PdfMerger()

    # Get all PDF files in the folder
    pdf_files = [file for file in os.listdir(folder_path) if file.endswith('.pdf')]

    # Merge all PDF files
    for file in pdf_files:
        file_path = os.path.join(folder_path, file)
        merger.append(file_path)

    # Write the merged PDF to the output file
    merger.write(output_file)

    # Close the merger
    merger.close()

# Example usage
# folder_path = 'path/to/folder'
# output_file = 'merged.pdf'
# merge_pdfs_in_folder(folder_path, output_file)


def save_attachments_from_subfolder(folder_name, save_folder, start_date=None, end_date=None):
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Get the root folder of the mailbox
    root_folder = namespace.GetDefaultFolder(6)  # 6 represents the Inbox folder

    # Get the subfolder by name
    subfolder = root_folder.Folders[folder_name]

    # Create the save folder if it doesn't exist
    os.makedirs(save_folder, exist_ok=True)

    # Iterate through the items in the subfolder
    for item in subfolder.Items:
        if start_date and item.ReceivedTime.date() < start_date:
            continue
        if end_date and item.ReceivedTime.date() > end_date:
            continue

        for attachment in item.Attachments:
            # Save each attachment to the specified folder
            attachment.SaveAsFile(os.path.join(save_folder, attachment.FileName))


def handle_save_attachments():
    subfolder_name = subfolder_entry.get()
    save_folder = save_folder_entry.get()
    start_date_str = start_date_entry.get()
    end_date_str = end_date_entry.get()

    # Validate inputs
    if not subfolder_name or not save_folder:
        messagebox.showerror("Error", "Please enter subfolder name and save folder path.")
        return

    try:
        start_date = datetime.datetime.strptime(start_date_str, "%m/%d/%Y").date() if start_date_str else None
        end_date = datetime.datetime.strptime(end_date_str, "%m/%d/%Y").date() if end_date_str else None
    except ValueError:
        messagebox.showerror("Error", "Invalid date format. Please use MM-DD-YYYY.")
        return

    # Save attachments from the subfolder within the specified period to the specified folder
    save_attachments_from_subfolder(subfolder_name, save_folder, start_date, end_date)
    messagebox.showinfo("Success", "Attachments saved successfully!")


# Create the Tkinter window
window = tk.Tk()
window.title("Save Attachments")
window.geometry("400x250")

# Subfolder name input
subfolder_label = tk.Label(window, text="Subfolder Name:")
subfolder_label.pack()
subfolder_entry = tk.Entry(window)
subfolder_entry.pack()

# Save folder path input
save_folder_label = tk.Label(window, text="Save Folder Path:")
save_folder_label.pack()
save_folder_entry = tk.Entry(window)
save_folder_entry.pack()

# Start date input
start_date_label = tk.Label(window, text="Start Date (MM-DD-YYYY):")
start_date_label.pack()
start_date_entry = tk.Entry(window)
start_date_entry.pack()

# End date input
end_date_label = tk.Label(window, text="End Date (MM-DD-YYYY):")
end_date_label.pack()
end_date_entry = tk.Entry(window)
end_date_entry.pack()

# Save button
save_button = tk.Button(window, text="Save Attachments", command=handle_save_attachments)
save_button.pack()

# Start the Tkinter event loop
window.mainloop()
