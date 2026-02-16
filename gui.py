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
import os
import queue
import threading
import subprocess
import customtkinter as ctk
from PIL import Image
from tkinter import filedialog
from main import main

# ----------------- DummyStream Fix for PyInstaller console=False -----------------
class DummyStream:
    def write(self, *args, **kwargs): pass
    def flush(self): pass

if sys.stderr is None:
    sys.stderr = DummyStream()
if sys.stdout is None:
    sys.stdout = DummyStream()

# ----------------- Config -----------------
program_name = "Rsg Doc remerge"
instructions_path = r'P:\Users\Justin\Program Shortcuts\Docs\PDF\RSG Document Merger.pdf'
input_type = "folder"  # "file", "files", "ofn", "folder"

# ----------------- Resource Path Helper -----------------
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# ----------------- GUI Class -----------------
class GUI:
    def __init__(self, master):
        self.master = master
        self.master.title(program_name)
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("dark-blue")
        master.configure(bg="#303030")

        # ----------------- Image Frame -----------------
        self.image_frame = ctk.CTkFrame(master, fg_color="#303030", corner_radius=0, width=270, height=130)
        self.image_frame.grid(row=0, column=0, rowspan=7, sticky="nsew")

        image_path = resource_path('assets/rvo_logo.png')
        pil_image = Image.open(image_path)
        self.image = ctk.CTkImage(light_image=pil_image, size=(270,130))
        self.image_label = ctk.CTkLabel(self.image_frame, image=self.image, text="", fg_color="#303030",
                                        corner_radius=0, width=270, height=130)
        self.image_label.pack(padx=20, pady=0, fill="both", expand=True)

        # Help Button
        self.help_dropdown = None
        self.help_button = ctk.CTkButton(self.image_frame, text="Help", width=30, height=30,
                                         command=self.show_help_dropdown, fg_color="#505050")
        self.help_button.place(x=5, y=5)

        # ----------------- Title -----------------
        text_color = "#8e979e" if ctk.get_appearance_mode() == "Dark" else "#808080"
        self.title_label = ctk.CTkLabel(master, text=program_name, font=("Times New Roman", 28), text_color=text_color)
        self.title_label.grid(row=0, column=1, columnspan=2, padx=0, pady=(100,30), sticky="new")

        # ----------------- Input Entry -----------------
        input_entry_placeholder = {
            "file": "Input File",
            "files": "Input File(s)",
            "ofn": "OFN List",
            "folder": "Input Folder"
        }.get(input_type, "Input")

        self.input_entry = ctk.CTkEntry(master, width=200, placeholder_text=input_entry_placeholder)
        self.input_entry.grid(row=1, column=1, sticky="ew", padx=(18,0), pady=(0,5))

        self.browse_input_button = ctk.CTkButton(master, text="Browse", command=self.browse_input, width=75)
        self.browse_input_button.grid(row=1, column=2, padx=(5,18), pady=(0,5), sticky="ew")

        # ----------------- TXT Entry (9-digit numbers) -----------------
        self.txt_entry = ctk.CTkEntry(master, width=200, placeholder_text="9-Digit TXT List")
        self.txt_entry.grid(row=2, column=1, sticky="ew", padx=(18,0), pady=(0,5))

        self.browse_txt_button = ctk.CTkButton(master, text="Browse TXT", command=self.browse_txt, width=75)
        self.browse_txt_button.grid(row=2, column=2, padx=(5,18), pady=(0,5), sticky="ew")

        # ----------------- Output Folder -----------------
        self.output_folder_entry = ctk.CTkEntry(master, width=200, placeholder_text="Output Folder")
        self.output_folder_entry.grid(row=3, column=1, sticky="ew", padx=(18,0), pady=(0,5))

        self.browse_output_button = ctk.CTkButton(master, text="Browse", command=self.browse_output, width=75)
        self.browse_output_button.grid(row=3, column=2, padx=(5,18), pady=(0,5), sticky="ew")

        # ----------------- Submit -----------------
        self.submit_button = ctk.CTkButton(master, text="Submit", command=self.submit_process)
        self.submit_button.grid(row=4, column=1, columnspan=2, padx=18, pady=(0,5), sticky="ew")

        # ----------------- Status -----------------
        self.processing_status_var = ctk.StringVar()
        self.processing_status_label = ctk.CTkLabel(master, textvariable=self.processing_status_var,
                                                    text_color="#808080", font=("Ariel", 16))
        self.processing_status_label.grid(row=5, column=1, columnspan=2, padx=0, pady=(35,39), sticky="ew")

        self.processing_dots = ["...", "   ", ".  ", ".. "]
        self.processing_dots_index = 0
        self.processing_dots_running = False

        # ----------------- Progress Bar -----------------
        progress_color = "#0091ff" if ctk.get_appearance_mode()=="Dark" else "#80c8ff"
        self.progress_bar = ctk.CTkProgressBar(master, width=200, height=10, corner_radius=4,
                                               progress_color=progress_color)
        self.progress_bar.grid(row=6, column=1, columnspan=2, padx=58, pady=(0,36), sticky="ew")
        self.progress_bar.set(0)
        self.progress_bar.grid_remove()

        self.progress_queue = queue.Queue()
        self.master.after(100, self.check_progress_queue)

        # ----------------- Grid Configuration -----------------
        for i in range(7):
            master.grid_rowconfigure(i, weight=0)
        for i in range(3):
            master.grid_columnconfigure(i, weight=0)

    # ----------------- Help Dropdown -----------------
    def show_help_dropdown(self):
        if self.help_dropdown and self.help_dropdown.winfo_exists():
            self.help_dropdown.destroy()
            self.master.unbind("<Button-1>")
            return

        self.help_dropdown = ctk.CTkFrame(self.image_frame, fg_color="#404040", corner_radius=6)
        self.help_dropdown.place(x=5, y=40)

        instructions_btn = ctk.CTkButton(self.help_dropdown, text="Instructions", width=120,
                                         command=self.show_instructions)
        instructions_btn.pack(pady=2, padx=2)

        self.master.bind("<Button-1>", self.on_click_outside_help, add="+")

    def on_click_outside_help(self, event):
        if self.help_dropdown and self.help_dropdown.winfo_exists():
            widget = event.widget
            if widget not in (self.help_dropdown, self.help_button) and not str(widget).startswith(str(self.help_dropdown)):
                self.help_dropdown.destroy()
                self.master.unbind("<Button-1>")

    def show_instructions(self):
        if os.path.exists(instructions_path):
            os.startfile(instructions_path)
        else:
            self.set_processing_status("Instructions document not found.")

    # ----------------- Processing Dots -----------------
    def animate_processing_dots(self):
        if self.processing_dots_running:
            dots = self.processing_dots[self.processing_dots_index]
            self.processing_status_var.set(f"Processing{dots}")
            self.processing_dots_index = (self.processing_dots_index + 1) % len(self.processing_dots)
            self.master.after(500, self.animate_processing_dots)

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

    # ----------------- Browse Functions -----------------
    def browse_input(self):
        input_path = ""
        if input_type == "files":
            input_paths = filedialog.askopenfilenames(filetypes=[("Excel files","*.xlsx *.xlsm *.xltx *.xltm")])
            if input_paths:
                self.input_entry.delete(0, "end")
                self.input_entry.insert(0, "|".join(input_paths))
        elif input_type == "ofn":
            input_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
            if input_path:
                with open(input_path, "r") as f:
                    lines = [l.strip() for l in f]
                self.input_entry.delete(0, "end")
                self.input_entry.insert(0, ",".join(lines))
        else:
            if input_type == "file":
                input_path = filedialog.askopenfilename()
            if input_type == "folder":
                input_path = filedialog.askdirectory()
            if input_path:
                self.input_entry.delete(0, "end")
                self.input_entry.insert(0, input_path)

    def browse_txt(self):
        input_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if not input_path:
            return
        try:
            with open(input_path, "r") as file:
                lines = file.readlines()
            nums = [line.strip() for line in lines if line.strip().isdigit() and len(line.strip()) == 9]
            if not nums:
                self.set_processing_status("No valid 9-digit numbers found.")
                return
            self.txt_entry.delete(0, "end")
            self.txt_entry.insert(0, ",".join(nums))
        except Exception as e:
            print("TXT read error:", e)
            self.set_processing_status("Error reading TXT file.")

    def browse_output(self):
        output_path = filedialog.askdirectory()
        if output_path:
            self.output_folder_entry.delete(0, "end")
            self.output_folder_entry.insert(0, output_path)

    # ----------------- Submit & Threading -----------------
    def submit_process(self):
        input_value = self.input_entry.get()
        txt_value = self.txt_entry.get()
        output_folder = self.output_folder_entry.get()
        if not output_folder or (not input_value and not txt_value):
            self.set_processing_status("Missing parameters.")
            return

        # Prioritize TXT input if provided
        if txt_value:
            input_value = [num.strip() for num in txt_value.split(",")]
        elif input_type in ["ofn","files"]:
            input_value = [num.strip() for num in input_value.replace("|",",").split(",")]

        threading.Thread(target=self.main_threaded, args=(input_value, output_folder)).start()
        self.submit_button.configure(state="disabled")
        self.set_processing_status("Processing...")

    def main_threaded(self, input_value, output_folder):
        try:
            self.progress_bar.set(0)
            main(input_value, output_folder, self.progress_queue)
            self.set_processing_status("Processing complete.")
        except Exception as e:
            print("Error during processing:", e)
            self.set_processing_status("Error during processing")
        finally:
            self.submit_button.configure(state="normal")

# ----------------- Main -----------------
if __name__ == "__main__":
    root = ctk.CTk()
    app = GUI(root)
    root.mainloop()