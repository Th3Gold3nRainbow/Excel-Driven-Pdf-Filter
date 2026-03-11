import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import os
import threading
import customtkinter as ctk # Modern UI Library
from tkinter import filedialog, messagebox

# Set the theme
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class Modern_PDF_Filter(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("EXPDF-Filter Pro")
        self.geometry("500x450")
        
        self.pdf_path = ""
        self.excel_path = ""
        self.df = None

        # --- UI ELEMENTS ---
        self.label_title = ctk.CTkLabel(self, text="EXPDF-Filter", font=("Roboto", 24, "bold"))
        self.label_title.pack(pady=20)

        # File Selection Frame
        self.frame = ctk.CTkFrame(self)
        self.frame.pack(pady=10, padx=20, fill="both", expand=True)

        self.btn_browse = ctk.CTkButton(self.frame, text="📁 SELECT PDF & EXCEL", 
                                        command=self.load_files, font=("Roboto", 14, "bold"))
        self.btn_browse.pack(pady=20)

        self.label_info = ctk.CTkLabel(self.frame, text="Status: Waiting for files...", text_color="gray")
        self.label_info.pack()

        # Column Selection
        ctk.CTkLabel(self.frame, text="Search Column:").pack(pady=(10, 0))
        self.column_selector = ctk.CTkComboBox(self.frame, values=["First select Excel"], width=250)
        self.column_selector.pack(pady=5)

        # Progress
        self.progress_label = ctk.CTkLabel(self, text="0%")
        self.progress_bar = ctk.CTkProgressBar(self, width=400)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=10)

        # Action Button
        self.btn_run = ctk.CTkButton(self, text="🚀 RUN EXTRACTION", 
                                     command=self.start_thread, fg_color="#2ecc71", hover_color="#27ae60",
                                     text_color="white", font=("Roboto", 16, "bold"))
        self.btn_run.pack(pady=20)

    def load_files(self):
        self.pdf_path = filedialog.askopenfilename(title="Select PDF", filetypes=[("PDF", "*.pdf")])
        self.excel_path = filedialog.askopenfilename(title="Select Excel", filetypes=[("Excel", "*.xlsx *.xls")])
        
        if self.pdf_path and self.excel_path:
            try:
                self.df = pd.read_excel(self.excel_path)
                self.column_selector.configure(values=list(self.df.columns))
                self.column_selector.set(self.df.columns[0])
                self.label_info.configure(text="Files ready ✅", text_color="#2ecc71")
            except Exception as e:
                messagebox.showerror("Error", f"Excel Error: {e}")

    def start_thread(self):
        threading.Thread(target=self.run_process, daemon=True).start()

    def run_process(self):
        col = self.column_selector.get()
        if not self.pdf_path or not self.df is not None:
            messagebox.showwarning("Missing Files", "Please select files first!")
            return

        try:
            self.btn_run.configure(state="disabled")
            search_vals = self.df[col].dropna().astype(str).tolist()
            reader = PdfReader(self.pdf_path)
            writer = PdfWriter()
            total = len(reader.pages)
            found = 0

            for i, page in enumerate(reader.pages):
                progress = (i + 1) / total
                self.progress_bar.set(progress)
                self.progress_label.configure(text=f"{int(progress*100)}% - Page {i+1}/{total}")
                
                text = page.extract_text().lower()
                if any(v.lower() in text for v in search_vals):
                    writer.add_page(page)
                    found += 1

            if found > 0:
                output = os.path.join(os.path.dirname(self.pdf_path), f"Extracted_{col}.pdf")
                with open(output, "wb") as f:
                    writer.write(f)
                messagebox.showinfo("Success", f"Done! {found} pages extracted.")
            else:
                messagebox.showwarning("No Match", "No matches found.")
        
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.btn_run.configure(state="normal")
            self.progress_bar.set(0)
            self.progress_label.configure(text="Ready")

if __name__ == "__main__":
    app = Modern_PDF_Filter()
    app.mainloop()
