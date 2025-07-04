import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk

class FileSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Splitter")
        self.root.geometry("450x460")
        self.root.resizable(False, False)

        try:
            self.root.iconphoto(False, tk.PhotoImage(file="icon.png"))
        except Exception:
            pass

        self.df = None
        self.file_path = ""

        try:
            logo = Image.open("icon.png").resize((50, 50))
            self.logo_img = ImageTk.PhotoImage(logo)
            tk.Label(root, image=self.logo_img).pack(pady=(15, 5))
        except Exception:
            pass

        tk.Label(root, text="Spreadsheet Splitter", font=("Helvetica", 16, "bold")).pack()

        tk.Button(root, text="Select Excel or CSV File", command=self.select_file,
                  height=2, width=30, font=("Segoe UI", 10, "bold")).pack(pady=10)

        tk.Label(root, text="Column to Split By:", font=("Segoe UI", 10)).pack()
        self.column_dropdown = ttk.Combobox(root, state="readonly", width=40)
        self.column_dropdown.pack(pady=5)

        tk.Label(root, text="Select Export Format:", font=("Segoe UI", 10)).pack()
        self.format_var = tk.StringVar(value="xlsx")
        self.format_dropdown = ttk.Combobox(root, textvariable=self.format_var,
                                            values=["xlsx", "csv"], state="readonly", width=10)
        self.format_dropdown.pack(pady=5)

        self.progress = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=300)
        self.progress.pack(pady=10)

        tk.Button(root, text="Split File", command=self.split_file,
                  height=2, width=20, font=("Segoe UI", 10, "bold"),
                  bg="#1d72e8", fg="white").pack(pady=10)

        self.status_label = tk.Label(root, text="", font=("Segoe UI", 10), fg="gray")
        self.status_label.pack()

        credit_frame = tk.Frame(root)
        credit_frame.pack(pady=(10, 5))
        tk.Label(credit_frame, text="By ", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        link = tk.Label(credit_frame, text="Aditya Kumar Choudhary", fg="blue", cursor="hand2",
                        font=("Segoe UI", 9, "underline"))
        link.pack(side=tk.LEFT)
        link.bind("<Button-1>", lambda e: self.open_github())

    def open_github(self):
        import webbrowser
        webbrowser.open_new("https://github.com/adi0905")

    def select_file(self):
        filetypes = [("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        self.file_path = filedialog.askopenfilename(title="Select your file", filetypes=filetypes)

        if not self.file_path:
            return

        try:
            if self.file_path.endswith(".csv"):
                try:
                    self.df = pd.read_csv(self.file_path, encoding="utf-8")
                except UnicodeDecodeError:
                    self.df = pd.read_csv(self.file_path, encoding="ISO-8859-1")
            else:
                self.df = pd.read_excel(self.file_path, engine="openpyxl")

            self.df.columns = [str(col).strip() for col in self.df.columns]
            self.column_dropdown["values"] = self.df.columns.tolist()
            self.column_dropdown.current(0)
            self.status_label.config(text="File loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file:\n{e}")
            self.status_label.config(text="File loading failed.")

    def split_file(self):
        if self.df is None:
            messagebox.showwarning("No File", "Please load a file first.")
            return

        split_column = self.column_dropdown.get()
        if not split_column:
            messagebox.showwarning("No Column", "Please select a column to split by.")
            return

        export_format = self.format_var.get()
        if export_format not in ["xlsx", "csv"]:
            messagebox.showwarning("Invalid Format", "Please choose xlsx or csv format.")
            return

        self.df[split_column] = self.df[split_column].astype(str).str.strip()
        unique_vals = self.df[split_column].dropna().unique()
        total = len(unique_vals)
        if total == 0:
            messagebox.showinfo("No Data", "No unique values found to split.")
            return

        output_dir = os.path.join(os.path.dirname(self.file_path), "Split_Output")
        os.makedirs(output_dir, exist_ok=True)

        self.progress["maximum"] = total
        self.progress["value"] = 0

        try:
            for i, val in enumerate(unique_vals, 1):
                safe_name = str(val).replace("/", "_").replace("\\", "_").replace(":", "_").replace("*", "_").replace("?", "_")
                filtered_df = self.df[self.df[split_column] == val].copy()
                filename = os.path.join(output_dir, f"{safe_name}.{export_format}")

                if export_format == "csv":
                    filtered_df.to_csv(filename, index=False)
                else:
                    filtered_df.to_excel(filename, index=False)

                self.progress["value"] = i
                self.root.update_idletasks()

            self.status_label.config(text=f"Exported {total} files to: {output_dir}")
            messagebox.showinfo("Done", f"{total} files exported to:\n{output_dir}")
        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred:\n{e}")
            self.status_label.config(text="Export failed.")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileSplitterApp(root)
    root.mainloop()
