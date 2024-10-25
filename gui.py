import os
import threading
import tkinter as tk
from ttkthemes import ThemedTk # pip install ttkthemes
from tkinter import filedialog, ttk, Menu
from tkinter.filedialog import askdirectory
from main import main

class GUI:
    def __init__(self, master):
        self.master = master
        self.master.title("RSG Document Merger")

        # Create a menu bar
        menu_bar = Menu(self.master)
        self.master.config(menu=menu_bar)

        # Create an Help menu
        instructions_menu = Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Help", menu=instructions_menu)

        # Add an Instructions option to the menu
        instructions_menu.add_command(label="Instructions", command=self.show_instructions)

        # Create labels and input widgets
        self.input_folder_label = tk.Label(master, text="Input Folder:")
        self.input_folder_label.grid(row=0, column=0, sticky=tk.E, padx=5, pady=(5,0))

        self.input_folder_var = tk.StringVar()
        self.input_folder_entry = ttk.Entry(master, textvariable=self.input_folder_var, width=30)
        self.input_folder_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=(5,0))

        self.browse_input_button = ttk.Button(master, text="Browse", command=self.browse_input)
        self.browse_input_button.grid(row=0, column=2, padx=10, pady=(5,0), sticky=tk.W)

        # Create labels and output widgets
        self.output_folder_label = tk.Label(master, text="Output Folder:")
        self.output_folder_label.grid(row=1, column=0, sticky=tk.E, padx=5, pady=5)

        self.output_folder_var = tk.StringVar()
        self.output_folder_entry = ttk.Entry(master, textvariable=self.output_folder_var, width=30)
        self.output_folder_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        self.browse_output_button = ttk.Button(master, text="Browse", command=self.browse_output)
        self.browse_output_button.grid(row=1, column=2, padx=10, pady=5, sticky=tk.W)

        # Create label for processing status
        self.processing_status_var = tk.StringVar()
        self.processing_status_label = tk.Label(master, textvariable=self.processing_status_var, fg="grey")
        self.processing_status_label.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky=tk.W)

        # Create generate button
        self.generate_button = ttk.Button(master, text="Merge PDFs", command=self.generate_excel)
        self.generate_button.grid(row=2, columnspan=3, padx=10, pady=10, sticky=tk.E)

    def set_processing_status(self, status):
        self.processing_status_var.set(status)

    def browse_input(self):
        folder_path = askdirectory()
        self.input_folder_var.set(folder_path)

    def browse_output(self):
        folder_path = filedialog.askdirectory()
        self.output_folder_var.set(folder_path)

    def open_file_explorer(self, folder_path):
        os.startfile(folder_path)

    def show_instructions(self):
        # Specify the path to your premade instructions document
        instructions_path = r'p:\Users\Justin\Program Shortcuts\Docs\PDF\RSG Document Merger.pdf'

        if os.path.exists(instructions_path):
            self.open_file_explorer(instructions_path)
        else:
            self.set_processing_status("Instructions document not found.")

    def generate_excel(self):
        input_folder = self.input_folder_var.get()
        output_folder = self.output_folder_var.get()

        if not input_folder or not output_folder:
            self.set_processing_status("Please select both input file and output folder.")
            return

        try:
            # Run the main function in a separate thread
            threading.Thread(target=self.main_threaded, args=(input_folder, os.path.abspath(self.output_folder_var.get()))).start()

            # Disable the Generate Excel button during processing
            self.generate_button.config(state=tk.DISABLED)

            # Set processing status
            self.set_processing_status("Processing...")

        except Exception as e:
            # Provide user-friendly error message
            print("Error:", e)
            self.set_processing_status("Error during processing")
            # Re-enable the Generate Excel button on error
            self.generate_button.config(state=tk.NORMAL)

    def main_threaded(self, input_folder, output_folder):
        try:
            main(input_folder, output_folder)
            # Provide user feedback upon completion
            self.set_processing_status("Files created successfuly.")
            # Open the output folder
            os.startfile(output_folder)
        except Exception as e:
            # Provide user-friendly error message
            print("Error during processing:", e)
            self.set_processing_status("Error during processing")
        finally:
            # Re-enable the Generate Excel button after processing
            self.generate_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    themed = ThemedTk(theme='plastik')
    app = GUI(themed)
    themed.mainloop()
