import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

class ExcelMergerApp:
    def __init__(self, master):
        self.master = master
        self.master.title("VanWinkle Merger")
        self.master.geometry("400x300")
        self.master.resizable(False, False)
        self.icon_path = "C:/99.Python/VanMerger.png"
        self.excel_icon_path = "C:/99.Python/Excel_icon.png"

        if os.path.exists(self.icon_path):
            self.master.iconphoto(True, tk.PhotoImage(file=self.icon_path))

        self.create_widgets()
        self.files = []

    def create_widgets(self):
        # Use themed style
        style = ttk.Style()
        style.theme_use('clam')

        self.main_frame = ttk.Frame(self.master, style="My.TFrame")
        self.main_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        self.file_frame = ttk.Frame(self.main_frame, style="My.TFrame")
        self.file_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        self.add_button = ttk.Button(self.main_frame, text="Add Excel Files", command=self.add_files, style="My.TButton")
        self.add_button.pack(pady=(0, 10), padx=10, fill=tk.X)

        self.output_folder_frame = ttk.Frame(self.main_frame, style="My.TFrame")
        self.output_folder_frame.pack(pady=(0, 10), padx=10, fill=tk.X)

        self.output_folder_label = ttk.Label(self.output_folder_frame, text="Output Folder:", style="My.TLabel")
        self.output_folder_label.pack(side=tk.LEFT, padx=(0, 5))

        self.output_folder_entry = ttk.Entry(self.output_folder_frame, width=30)
        self.output_folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        self.output_folder_button = ttk.Button(self.output_folder_frame, text="Browse", command=self.select_output_folder, style="My.TButton")
        self.output_folder_button.pack(side=tk.LEFT)

        self.output_filename_frame = ttk.Frame(self.main_frame, style="My.TFrame")
        self.output_filename_frame.pack(pady=(0, 10), padx=10, fill=tk.X)

        self.output_filename_label = ttk.Label(self.output_filename_frame, text="Output Filename:", style="My.TLabel")
        self.output_filename_label.pack(side=tk.LEFT, padx=(0, 5))

        self.output_filename_entry = ttk.Entry(self.output_filename_frame, width=30)
        self.output_filename_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.output_filename_entry.insert(tk.END, "combined_excel_file.xlsx")

        self.merge_button = ttk.Button(self.main_frame, text="Merge Excel Files", command=self.merge_files, style="My.TButton")
        self.merge_button.pack(pady=(0, 10), padx=10, fill=tk.X)

        # Load the Excel icon image with smaller size
        if os.path.exists(self.excel_icon_path):
            self.load_icons(30, 30)  # Adjust the size as needed

        # Configure style
        style.configure('My.TFrame', background='white', borderwidth=0)
        style.configure('My.TLabel', background='white')
        style.configure('My.TButton', background='#4CAF50', foreground='white', borderwidth=0, relief="raised", font=('Helvetica', 12, 'bold'))

    def load_icons(self, icon_width, icon_height):
        # Load the Excel icon image
        self.excel_icon = tk.PhotoImage(file=self.excel_icon_path)

        # Resize the icon to the specified size
        self.excel_icon = self.excel_icon.subsample(self.excel_icon.width() // icon_width, self.excel_icon.height() // icon_height)

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        for file in files:
            filename = os.path.basename(file)
            label = ttk.Label(self.file_frame, text=filename, image=self.excel_icon, compound=tk.LEFT, style="My.TLabel")
            label.pack(fill=tk.X)
            self.files.append(file)

    def select_output_folder(self):
        folder = filedialog.askdirectory()
        self.output_folder_entry.delete(0, tk.END)
        self.output_folder_entry.insert(0, folder)

    def merge_files(self):
        output_folder = self.output_folder_entry.get()
        if not output_folder:
            messagebox.showerror("Error", "Please select an output folder")
            return

        if not self.files:
            messagebox.showerror("Error", "Please add at least one Excel file")
            return

        output_filename = self.output_filename_entry.get().strip()
        if not output_filename:
            messagebox.showerror("Error", "Please enter a valid output filename")
            return

        if not output_filename.endswith('.xlsx'):
            output_filename += '.xlsx'

        dfs = [pd.read_excel(file_path) for file_path in self.files]
        combined_df = pd.concat(dfs, ignore_index=True)
        output_file_path = os.path.join(output_folder, output_filename)
        combined_df.to_excel(output_file_path, index=False)
        messagebox.showinfo("Success you legend !", f"Merged file saved at {output_file_path}")

        # Reset the application state
        self.reset_state()

    def reset_state(self):
        self.files = []
        self.output_folder_entry.delete(0, tk.END)
        self.output_filename_entry.delete(0, tk.END)
        self.output_filename_entry.insert(tk.END, "combined_excel_file.xlsx")
        for widget in self.file_frame.winfo_children():
            widget.destroy()

def main():
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
