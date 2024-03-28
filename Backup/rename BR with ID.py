import os
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
import re

def select_folder():
    folder_path = filedialog.askdirectory(title="Select Folder")
    if folder_path:
        entry_folder.delete(0, tk.END)
        entry_folder.insert(tk.END, folder_path)

def select_output_folder():
    output_folder = filedialog.askdirectory(title="Select Output Folder")
    if output_folder:
        entry_output.delete(0, tk.END)
        entry_output.insert(tk.END, output_folder)

def rename_files_with_id():
    folder_path = entry_folder.get()
    output_folder = entry_output.get()

    if not folder_path or not output_folder:
        messagebox.showerror("Error", "Please select input and output folders.")
        return

    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Loop through files in the input folder
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path):
            # Open the Word file
            doc = Document(file_path)
            # Read the content of the Word file
            content = ""
            for paragraph in doc.paragraphs:
                content += paragraph.text
            # Find the 17-18 ID digits
            matches = re.findall(r"[0-9A-Za-z]{17,18}", content)
            if matches:
                id = matches[0]
                # Rename the file using the ID
                new_name = id + os.path.splitext(filename)[1]
                new_path = os.path.join(output_folder, new_name)
                # Handle duplicate file names
                count = 1
                while os.path.exists(new_path):
                    base_name, extension = os.path.splitext(new_name)
                    new_name = f"{base_name}_{count}{extension}"
                    new_path = os.path.join(output_folder, new_name)
                    count += 1
                # Rename the file
                if file_path != new_path:
                    os.rename(file_path, new_path)

    messagebox.showinfo("Renaming Complete", "Files renamed successfully.")

# Create the Tkinter window
window = tk.Tk()
window.title("File Renamer")

# Create and place the input folder selection widgets
label_folder = tk.Label(window, text="Select Folder:")
label_folder.pack()

entry_folder = tk.Entry(window, width=50)
entry_folder.pack()

button_select_folder = tk.Button(window, text="Select", command=select_folder)
button_select_folder.pack()

# Create and place the output folder selection widgets
label_output = tk.Label(window, text="Select Output Folder:")
label_output.pack()

entry_output = tk.Entry(window, width=50)
entry_output.pack()

button_select_output = tk.Button(window, text="Select", command=select_output_folder)
button_select_output.pack()

# Create the rename button
button_rename = tk.Button(window, text="Rename Files", command=rename_files_with_id)
button_rename.pack()

# Start the Tkinter event loop
window.mainloop()
