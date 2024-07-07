import pandas as pd
import os
import customtkinter as ctk
from tkinter import filedialog, messagebox, StringVar, END
import webbrowser

def select_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[
        ("Excel files", "*.xlsx;*.xls"),
        ("CSS files", "*.css"),
        ("Text files", "*.txt"),
        ("All files", "*.*")
    ])
    excel_path_var.set(file_path)

def select_folder():
    folder_path = filedialog.askdirectory()
    folder_path_var.set(folder_path)

def rename_files():
    excel_path = excel_path_var.get()
    folder_path = folder_path_var.get()
    start_row = int(start_row_var.get())
    end_row = int(end_row_var.get())

    if not excel_path or not folder_path:
        messagebox.showerror("Error", "Please select both the Excel file and the folder.")
        return

    try:
        df = pd.read_excel(excel_path)

        for index, row in df.iloc[start_row-1:end_row].iterrows():
            old_name = str(row['A'])
            new_name = str(row['B'])

            old_file_path = None
            for file in os.listdir(folder_path):
                if file.startswith(old_name):
                    old_file_path = os.path.join(folder_path, file)
                    break
            
            if old_file_path:
                # Extract the file extension
                _, file_extension = os.path.splitext(old_file_path)
                new_file_path = os.path.join(folder_path, new_name + file_extension)
                
                if not os.path.exists(new_file_path):
                    os.rename(old_file_path, new_file_path)
                    log_text.insert(END, f'Renamed: {old_name} to {new_name}\n')
                else:
                    log_text.insert(END, f'File already exists: {new_name}\n')
            else:
                log_text.insert(END, f'File not found: {old_name}\n')

        messagebox.showinfo("Success", "Files renamed successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def toggle_theme():
    if dark_light_mode_var.get() == "0":
        ctk.set_appearance_mode("Dark")
    else:
        ctk.set_appearance_mode("Light")

def clear_text():
    log_text.delete("1.0", END)  # Delete all text in the Text widget

def open_template_url(event):
    webbrowser.open_new("./template.xlsx")  # Replace with your actual URL

# Basic parameters and initializations
ctk.set_appearance_mode("Dark")  # Supported modes: Light, Dark, System
ctk.set_default_color_theme("green")  # Supported themes: green, dark-blue, blue

# App Class
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Bulk File Renamer")
        self.geometry("660x600")  # Set the initial size of the window
        self.wm_minsize(660, 300)

        # Define Tkinter variables
        global excel_path_var, folder_path_var, start_row_var, end_row_var, dark_light_mode_var, log_text
        excel_path_var = StringVar()
        folder_path_var = StringVar()
        start_row_var = StringVar(value="1")
        end_row_var = StringVar(value="100")
        dark_light_mode_var = StringVar(value="1")  # 1 for Light mode, 0 for Dark mode

        # Configure the grid layout
        self.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)
        self.grid_rowconfigure(3, weight=1)

        # Create and place widgets with modern look
        font = ("Segoe UI", 12)

        self.label_excel = ctk.CTkLabel(self, text="Source File:", font=font)
        self.label_excel.grid(row=0, column=0, padx=10, pady=10, sticky="e")

        self.entry_excel = ctk.CTkEntry(self, textvariable=excel_path_var, font=font)
        self.entry_excel.grid(row=0, column=1,columnspan=3, padx=10, pady=10, sticky="ew")

        self.button_browse_excel = ctk.CTkButton(self, text="Browse", command=select_excel_file)
        self.button_browse_excel.grid(row=0, column=4, padx=10, pady=10, sticky="w")

        self.label_folder = ctk.CTkLabel(self, text="Select Folder:", font=font)
        self.label_folder.grid(row=1, column=0, padx=10, pady=10, sticky="e")

        self.entry_folder = ctk.CTkEntry(self, textvariable=folder_path_var, font=font)
        self.entry_folder.grid(row=1, column=1,columnspan=3, padx=10, pady=10, sticky="ew")

        self.button_browse_folder = ctk.CTkButton(self, text="Browse", command=select_folder)
        self.button_browse_folder.grid(row=1, column=4, padx=10, pady=10, sticky="w")

        self.label_start_row = ctk.CTkLabel(self, text="Start Row:", font=font)
        self.label_start_row.grid(row=2, column=0, padx=10, pady=10, sticky="e")

        self.entry_start_row = ctk.CTkEntry(self, textvariable=start_row_var, font=font)
        self.entry_start_row.grid(row=2, column=1, padx=10, pady=10, sticky="w")

        self.label_end_row = ctk.CTkLabel(self, text="End Row:", font=font)
        self.label_end_row.grid(row=2, column=2, padx=10, pady=10, sticky="e")

        self.entry_end_row = ctk.CTkEntry(self, textvariable=end_row_var, font=font)
        self.entry_end_row.grid(row=2, column=3, padx=10, pady=10, sticky="e")

        log_text = ctk.CTkTextbox(self, font=font)
        log_text.grid(row=3, column=0, columnspan=5, sticky="nsew",padx=10, pady=10)

        # Dark/Light mode toggle button
        self.dark_light_mode_button = ctk.CTkCheckBox(self, text="Dark Mode", variable=dark_light_mode_var, command=toggle_theme)
        self.dark_light_mode_button.grid(row=4, column=0, padx=10, pady=10, sticky="e")

        self.button_start = ctk.CTkButton(self, text="START", command=rename_files)
        self.button_start.grid(row=4, column=1, padx=10, pady=10, sticky="w")

        self.button_clear = ctk.CTkButton(self, text="CLEAR", command=clear_text)
        self.button_clear.grid(row=4, column=3, padx=10, pady=10, sticky="w")

        self.label_crate = ctk.CTkLabel(self, text="CREATED BY ADILBHAGAT", font=font, text_color="red")
        self.label_crate.grid(row=4, column=4, padx=10, pady=10, sticky="w")

        self.label_note = ctk.CTkLabel(self, text="Note :- First row add A(oldname) B(Newname)", font=font, text_color="red")
        self.label_note.grid(row=5, column=0, columnspan=5, sticky="we")

        self.label_template = ctk.CTkLabel(self, text="For a template file, please visit the application directory and look for 'template.xlsx'.", font=font, text_color="green",cursor="hand2")
        self.label_template.grid(row=6, column=0, columnspan=5, sticky="we")
        self.label_template.bind("<Button-1>", open_template_url)


if __name__ == "__main__":
    app = App()
    app.mainloop()
