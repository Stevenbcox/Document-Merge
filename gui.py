# Compile using the following command:
# pyinstaller --onefile --noconsole --add-data ".\assets\rvo_logo.png;assets" .\gui.py

# Add the following code to the main.py file to handle progress updates:
# Import the necessary modules
# from utils import progress_callback

# main function
# (add the progress_queue=None parameter)
# (adjust for loop to include the index and range(total_files))

# Add to the end of the for loop in the main function:
# process = (index + 1) / total_files
# progress_callback(progress_queue, process)

import sys

# Fixes console=False issue in PyInstaller
class DummyStream:
    def write(self, *args, **kwargs): pass
    def flush(self): pass

if sys.stderr is None:
    sys.stderr = DummyStream()
if sys.stdout is None:
    sys.stdout = DummyStream()

import os
import queue
import threading
import subprocess
import customtkinter as ctk
from PIL import Image
from tkinter import  filedialog
from main import main

# Define the program name and input type
program_name = "RSG Document Merger"  # Change this to your program name
instructions_path = r'P:\Users\Justin\Program Shortcuts\Docs\PDF\RSG Document Merger.pdf' # Path to the instructions
input_type = "folder"  # Type file, files, ofn, or folder depending on input type

# Function to get the absolute path to a resource
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

class GUI:
    def __init__(self, master):
        self.master = master
        self.master.title(f"{program_name}")

        # Set appearance mode and color theme
        ctk.set_appearance_mode("System")  # or "Dark", "Light"

        ctk.set_default_color_theme("dark-blue")
        
        master.configure(bg="#303030")

        self.image_frame = ctk.CTkFrame(master, fg_color="#303030", corner_radius=0, width=270, height=130)
        self.image_frame.grid(row=0, column=0, rowspan=6, sticky="nsew")

        # Use this for your image path
        image_path = resource_path('assets/rvo_logo.png')
        pil_image = Image.open(image_path)

        # Pass the PIL image to CTkImage
        self.image = ctk.CTkImage(light_image=pil_image, size=(270, 130))
        self.image_label = ctk.CTkLabel(
            self.image_frame,
            image=self.image,
            text="",
            fg_color="#303030",
            corner_radius=0,
            width=270,
            height=130
        )
        self.image_label.pack(padx=20, pady=0, fill="both", expand=True)

        # Configure the text color if the appearance mode is dark
        if ctk.get_appearance_mode() == "Dark":
            text_color = "#8e979e"
            progress_bar_color = "#80c8ff"
        else:
            text_color = "#808080"

        # Content column
        self.title_label = ctk.CTkLabel(
            master,
            text=f"{program_name}",
            font=("Times New Roman", 28),
            text_color=text_color
        )
        self.title_label.grid(row=0, column=1, columnspan=2, padx=0, pady=(100,30), sticky="new")

        # Special handling for radio_4 as a checkbox
        input_entry_placeholder = ""
        if input_type == "file":
            input_entry_placeholder = "Input File"
        elif input_type == "files":
            input_entry_placeholder = "Input File(s)"
        elif input_type == "ofn":
            input_entry_placeholder = "OFN List"
        elif input_type == "folder":
            input_entry_placeholder = "Input Folder"

        self.input_entry = ctk.CTkEntry(
            master,
            width=200,
            placeholder_text=f"{input_entry_placeholder}",
        )
        self.input_entry.grid(row=1, column=1, sticky="ew", padx=(18,0), pady=(0,5))

        self.browse_input_button = ctk.CTkButton(master, text="Browse", command=self.browse_input, width=75)
        self.browse_input_button.grid(row=1, column=2, padx=(5,18), pady=(0,5), sticky="ew")

        self.output_folder_entry = ctk.CTkEntry(
            master,
            width=200,
            placeholder_text="Output Folder",
        )
        self.output_folder_entry.grid(row=2, column=1, sticky="ew", padx=(18,0), pady=(0,5))

        self.browse_output_button = ctk.CTkButton(master, text="Browse", command=self.browse_output, width=75)
        self.browse_output_button.grid(row=2, column=2, padx=(5,18), pady=(0,5), sticky="ew")

        # Submit button
        self.submit_button = ctk.CTkButton(master, text="Submit", command=self.submit_process)
        self.submit_button.grid(row=3, column=1, columnspan=2, padx=18, pady=(0,5), sticky="ew")

        # Processing status
        self.processing_status_var = ctk.StringVar()
        self.processing_status_label = ctk.CTkLabel(
            master,
            textvariable=self.processing_status_var,
            text_color="#808080",
            font=("Ariel", 16)
        )
        self.processing_status_label.grid(row=4, column=1, columnspan=2, padx=0, pady=(35,39), sticky="ew")
        # self.processing_status_label.configure(anchor="center")

        # In the GUI class __init__ method, add:
        self.processing_dots = ["...", "   ", ".  ", ".. "]
        self.processing_dots_index = 0
        self.processing_dots_running = False

        if ctk.get_appearance_mode() == "Dark":
            self.progress_bar = ctk.CTkProgressBar(master, width=200, height=10, corner_radius=4, progress_color="#0091ff")
        else:
            self.progress_bar = ctk.CTkProgressBar(master, width=200, height=10, corner_radius=4)

        self.progress_bar.grid(row=5, column=1, columnspan=2, padx=58, pady=(0,36), sticky="ew")
        self.progress_bar.set(0)

        self.progress_queue = queue.Queue()
        self.master.after(100, self.check_progress_queue)

        self.progress_bar.grid_remove()

        # Configure grid weights
        for i in range(5):
            master.grid_rowconfigure(i, weight=0)

        for i in range(3):
            master.grid_columnconfigure(i, weight=0)

        self.help_dropdown = None
        self.help_button = ctk.CTkButton(
            self.image_frame,
            text="Help",
            width=30,
            height=30,
            command=self.show_help_dropdown,
            fg_color="#505050"
        )
        self.help_button.place(x=5, y=5)

    def show_help_dropdown(self):
        if self.help_dropdown and self.help_dropdown.winfo_exists():
            self.help_dropdown.destroy()
            self.master.unbind("<Button-1>")
            return

        self.help_dropdown = ctk.CTkFrame(self.image_frame, fg_color="#404040", corner_radius=6)
        self.help_dropdown.place(x=5, y=40)

        instructions_btn = ctk.CTkButton(
            self.help_dropdown, text="Instructions", width=120, command=self.show_instructions
        )
        instructions_btn.pack(pady=2, padx=2)

        # Bind a global click event
        self.master.bind("<Button-1>", self.on_click_outside_help, add="+")

    def on_click_outside_help(self, event):
        if self.help_dropdown and self.help_dropdown.winfo_exists():
            widget = event.widget
            # Check if click is outside the dropdown and help button
            if widget not in (self.help_dropdown, self.help_button) and not str(widget).startswith(
                    str(self.help_dropdown)):
                self.help_dropdown.destroy()
                self.master.unbind("<Button-1>")

    def open_edit_docs(self):
        # Specify the path to your premade instructions document
        edit_docs_path = r'P:\Users\Justin\Projects\efile_contested_docs\document_names.json'

        if os.path.exists(edit_docs_path):
            os.startfile(edit_docs_path)
        else:
            self.set_processing_status("Document Names file not found.")

    def show_instructions(self):
        if os.path.exists(instructions_path):
            os.startfile(instructions_path)
        else:
            self.set_processing_status("Instructions document not found.")

    # Animate processing dots:
    def animate_processing_dots(self):
        if self.processing_dots_running:
            dots = self.processing_dots[self.processing_dots_index]
            self.processing_status_var.set(f"Processing{dots}")
            self.processing_dots_index = (self.processing_dots_index + 1) % len(self.processing_dots)
            self.master.after(500, self.animate_processing_dots)

    # Update set_processing_status method:
    def set_processing_status(self, status):
        if status == "Processing...":
            self.progress_bar.grid()
            self.processing_status_label.grid_configure(pady=(35, 0))
            self.processing_dots_running = True
            self.processing_dots_index = 0
            self.animate_processing_dots()
        else:
            self.progress_bar.grid_remove()
            self.processing_status_label.grid_configure(pady=(35, 58))
            self.processing_dots_running = False
            self.processing_status_var.set(status)

    def check_progress_queue(self):
        try:
            while True:
                value = self.progress_queue.get_nowait()
                self.progress_bar.set(value)
        except queue.Empty:
            pass
        self.master.after(100, self.check_progress_queue)

    def browse_input(self):
        input_path = ""

        # Open the file explorer to select a file/folder
        if input_type == "files":
            input_paths = filedialog.askopenfilenames(
                filetypes=[("Excel files", "*.xlsx *.xlsm *.xltx *.xltm")] # Adjust file types as needed
            )
            if input_paths:
                # Join file paths into a space-separated string for the entry widget
                input_paths_str = '|'.join(input_paths)
                self.input_entry.delete(0, "end")
                self.input_entry.insert(0, input_paths_str)
        elif input_type == "ofn":
            # Look for text files only
            input_path = filedialog.askopenfilename(
                filetypes=[("Text files", "*.txt")]
            )
            # Create a list from the file content
            with open(input_path, 'r') as file:
                file_numbers = [line.strip() for line in file]  # Create a list
            # Put the list into the input entry
            self.input_entry.delete(0, "end")
            self.input_entry.insert(0, ','.join(file_numbers))  # Join the list with a comma
        else:
            if input_type == "file":
                input_path = filedialog.askopenfilename()
            if input_type == "folder":
                input_path = filedialog.askdirectory()

            if input_path:
                self.input_entry.delete(0, "end")
                self.input_entry.insert(0, input_path)

    def browse_output(self):
        # Open the file explorer to select a folder
        output_path = filedialog.askdirectory()
        if output_path:
            self.output_folder_entry.delete(0, "end")
            self.output_folder_entry.insert(0, output_path)

    def run_merge(self):
        program_path = r"P:\Users\Justin\Projects\excel_merge\dist\gui\gui.exe"
        subprocess.Popen(program_path, creationflags=subprocess.CREATE_NO_WINDOW)

    def run_format(self):
        program_path = r"P:\Users\Justin\Projects\tax_disc_split_tool\dist\gui\gui.exe"
        subprocess.Popen(program_path, creationflags=subprocess.CREATE_NO_WINDOW)

    def submit_process(self):
        input_value = self.input_entry.get()
        if not input_value:
            self.set_processing_status("Missing parameters.")
            return
        output_folder = self.output_folder_entry.get()
        if not output_folder:
            self.set_processing_status("Missing parameters.")
            return

        try:
            # Run the main function in a separate thread
            threading.Thread(target=self.main_threaded, args=(input_value, output_folder)).start()

            # Disable the submit Excel button during processing
            self.submit_button.configure(state="disabled")

            # Set processing status
            self.set_processing_status("Processing...")

        except Exception as e:
            # Provide user-friendly error message
            print("Error:", e)
            self.set_processing_status("Error during processing")

            # Re-enable the submit Excel button on error
            self.submit_button.configure(state="normal")

    def main_threaded(self, input_value, output_folder):
        try:
            # Reset the progress bar
            self.progress_bar.set(0)

            if input_type == "ofn":
                input_value = [num.strip() for num in input_value.split(',')]
            if input_type == "files":
                input_value = input_value.split('|')

            main(input_value, output_folder, self.progress_queue)

            # Provide user feedback upon completion
            self.set_processing_status("Processing complete.")
        except Exception as e:
            # Provide user-friendly error message
            print("Error during processing:", e)
        finally:
            # Re-enable the submit Excel button after processing
            self.submit_button.configure(state="normal")

if __name__ == "__main__":
    root = ctk.CTk()
    app = GUI(root)
    root.mainloop()
