import tkinter as tk
from tkinter import filedialog

def select_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        label_file.config(text=f"Selected: {file_path}")
        selected_file.set(file_path)

def select_directory():
    output_directory = filedialog.askdirectory(title="Select Output Directory")
    if output_directory:
        label_dir.config(text=f"Selected: {output_directory}")
        selected_directory.set(output_directory)

def option_selected(choice):
    label_option.config(text=f"Selected option: {choice}")

def submit_action():
    file = selected_file.get()
    directory = selected_directory.get()
    option = selected_option.get()
    submit_label.config(text=f"File: {file}, Directory: {directory}, Option: {option}")

# Create the root window
root = tk.Tk()
root.title("AutoDocs")

# Set the window size
root.geometry("600x500")
root.config(bg="#f0f0f0")  # Background color

# Variables to store user selections
selected_file = tk.StringVar(root, value="No file selected")
selected_directory = tk.StringVar(root, value="No directory selected")

# Define font and styling options
font_label = ("Helvetica", 12)
font_button = ("Helvetica", 12, "bold")

# Create a frame for file selection
frame_file = tk.Frame(root, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
frame_file.pack(padx=20, pady=20, fill="x")

label_file_title = tk.Label(frame_file, text="File Selection", font=font_label, bg="#e6e6e6")
label_file_title.pack(anchor="w")

button_file = tk.Button(frame_file, text="Select File", command=select_file, font=font_button, bg="#4CAF50", fg="white")
button_file.pack(side="left", padx=10, pady=5)

label_file = tk.Label(frame_file, textvariable=selected_file, bg="#e6e6e6", font=font_label)
label_file.pack(side="left", padx=10)

# Create a frame for directory selection
frame_dir = tk.Frame(root, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
frame_dir.pack(padx=20, pady=20, fill="x")

label_dir_title = tk.Label(frame_dir, text="Directory Selection", font=font_label, bg="#e6e6e6")
label_dir_title.pack(anchor="w")

button_dir = tk.Button(frame_dir, text="Select Output Directory", command=select_directory, font=font_button, bg="#4CAF50", fg="white")
button_dir.pack(side="left", padx=10, pady=5)

label_dir = tk.Label(frame_dir, textvariable=selected_directory, bg="#e6e6e6", font=font_label)
label_dir.pack(side="left", padx=10)

# Create a frame for the option menu
frame_option = tk.Frame(root, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
frame_option.pack(padx=20, pady=20, fill="x")

label_option_title = tk.Label(frame_option, text="Select an Option", font=font_label, bg="#e6e6e6")
label_option_title.pack(anchor="w")

options = ["Option 1", "Option 2", "Option 3", "Option 4"]
selected_option = tk.StringVar(root)
selected_option.set(options[0])

option_menu = tk.OptionMenu(frame_option, selected_option, *options, command=option_selected)
option_menu.config(font=font_button, bg="#4CAF50", fg="white")
option_menu.pack(side="left", padx=10, pady=5)

label_option = tk.Label(frame_option, text="No option selected", bg="#e6e6e6", font=font_label)
label_option.pack(side="left", padx=10)

# Create a Submit button
submit_button = tk.Button(root, text="Submit", command=submit_action, font=font_button, bg="#008CBA", fg="white")
submit_button.pack(pady=20)

# Label to show the result after submission
submit_label = tk.Label(root, text="", bg="#f0f0f0", font=font_label)
submit_label.pack(pady=10)

# Run the application
root.mainloop()
