import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import win32com.client
except ImportError:
    messagebox.showerror(
        "Missing Library",
        "pywin32 is required.\nRun: pip install pywin32"
    )
    raise SystemExit

# File type configuration
SUPPORTED_FORMATS = {
    # Word
    ".docx": "Word",
    ".doc": "Word",
    ".docm": "Word",
    ".dotx": "Word",
    ".dotm": "Word",
    ".rtf": "Word",
    ".odt": "Word",
    # Excel
    ".xlsx": "Excel",
    ".xls": "Excel",
    ".xlsm": "Excel",
    ".xltx": "Excel",
    ".xltm": "Excel",
    ".csv": "Excel",
    # PowerPoint
    ".pptx": "PowerPoint",
    ".ppt": "PowerPoint",
    ".pptm": "PowerPoint",
    ".potx": "PowerPoint",
    ".potm": "PowerPoint",
}

PDF_EXPORT_FORMATS = {
    "Word": 17,        # wdExportFormatPDF
    "Excel": 0,         # xlTypePDF
    "PowerPoint": 32,   # ppSaveAsPDF
}


def convert_office_to_pdf_safe(input_path, output_path, office_app):
    """
    Convert any supported Office file to PDF using the appropriate app.
    Works even if the file is open elsewhere.
    """
    app = None
    doc = None
    try:
        # Launch the right Office application
        app = win32com.client.Dispatch(f"{office_app}.Application")
        app.Visible = False

        if office_app == "Word":
            doc = app.Documents.Open(input_path, ReadOnly=True)
            doc.ExportAsFixedFormat(output_path, ExportFormat=17)

        elif office_app == "Excel":
            doc = app.Workbooks.Open(input_path, ReadOnly=True)
            doc.ExportAsFixedFormat(0, output_path)  # 0 = xlTypePDF

        elif office_app == "PowerPoint":
            doc = app.Presentations.Open(input_path, ReadOnly=True)
            doc.ExportAsFixedFormat(output_path, 32)  # 32 = ppSaveAsPDF

        return True

    except Exception as e:
        raise e

    finally:
        # Clean up
        if doc:
            doc.Close(SaveChanges=False)
        if app:
            app.Quit()


class OfficeToPDFConverter:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Family PDF Tool")
        self.window.geometry("560x320")
        self.window.resizable(False, False)

        default_font = ("Segoe UI", 10)
        self.window.option_add("*Font", default_font)

        self.filepath = tk.StringVar()
        self.filename_display = tk.StringVar(value="No file selected")
        self.filetype_display = tk.StringVar(value="")

        self._converting = False

        self._setup_ui()

    def _setup_ui(self):
        main_frame = tk.Frame(self.window, padx=30, pady=20)
        main_frame.pack(expand=True, fill="both")

        # Title
        title = tk.Label(
            main_frame,
            text="Office to PDF Converter",
            font=("Segoe UI", 14, "bold"),
        )
        title.pack(pady=(0, 5))

        # Supported formats hint
        subtitle = tk.Label(
            main_frame,
            text="Word • Excel • PowerPoint",
            font=("Segoe UI", 9),
            fg="gray",
        )
        subtitle.pack(pady=(0, 15))

        # File info display
        self.file_label = tk.Label(
            main_frame,
            textvariable=self.filename_display,
            wraplength=460,
            anchor="center",
            justify="center",
        )
        self.file_label.pack(pady=(0, 5))

        self.type_label = tk.Label(
            main_frame,
            textvariable=self.filetype_display,
            fg="#0078D4",
            font=("Segoe UI", 9, "bold"),
        )
        self.type_label.pack(pady=(0, 5))

        self.status_var = tk.StringVar()
        self.status_label = tk.Label(
            main_frame, textvariable=self.status_var, fg="gray"
        )
        self.status_label.pack(pady=(0, 5))

        self.progress = ttk.Progressbar(
            main_frame, mode="indeterminate", length=260
        )

        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(pady=10)

        self.select_btn = tk.Button(
            btn_frame, text="Select File", width=15, command=self.select_file
        )
        self.select_btn.grid(row=0, column=0, padx=5)

        self.convert_btn = tk.Button(
            btn_frame,
            text="Convert to PDF",
            width=15,
            command=self.start_conversion,
            state="disabled",
        )
        self.convert_btn.grid(row=0, column=1, padx=5)

    def select_file(self):
        # Build file type filter for dialog
        filetypes = [
            ("All Supported Files", "*.docx;*.doc;*.rtf;*.odt;*.xlsx;*.xls;*.csv;*.pptx;*.ppt"),
            ("Word Documents", "*.docx;*.doc;*.docm;*.dotx;*.dotm;*.rtf;*.odt"),
            ("Excel Spreadsheets", "*.xlsx;*.xls;*.xlsm;*.csv"),
            ("PowerPoint Presentations", "*.pptx;*.ppt;*.pptm"),
            ("All Files", "*.*"),
        ]

        path = filedialog.askopenfilename(
            title="Select an Office Document",
            filetypes=filetypes,
        )
        if path:
            ext = os.path.splitext(path)[1].lower()
            office_app = SUPPORTED_FORMATS.get(ext)

            if office_app:
                self.filepath.set(path)
                self.filename_display.set(os.path.basename(path))
                self.filetype_display.set(f"Detected: {office_app} document")
                self.convert_btn.config(state="normal")
            else:
                messagebox.showwarning(
                    "Unsupported Format",
                    f"The file type '{ext}' is not supported.\n\n"
                    "Supported: Word, Excel, and PowerPoint documents."
                )
        else:
            self.filepath.set("")
            self.filename_display.set("No file selected")
            self.filetype_display.set("")
            self.convert_btn.config(state="disabled")

    def start_conversion(self):
        if self._converting:
            return

        input_path = self.filepath.get()
        if not input_path:
            return

        output_path = os.path.splitext(input_path)[0] + ".pdf"
        if os.path.exists(output_path):
            overwrite = messagebox.askyesno(
                "File Exists",
                f"{os.path.basename(output_path)}\nalready exists.\nOverwrite?",
            )
            if not overwrite:
                return

        ext = os.path.splitext(input_path)[1].lower()
        office_app = SUPPORTED_FORMATS[ext]

        self._converting = True
        self._set_ui_state(converting=True)

        thread = threading.Thread(
            target=self._convert_worker,
            args=(input_path, output_path, office_app),
            daemon=True,
        )
        thread.start()

    def _convert_worker(self, input_path, output_path, office_app):
        try:
            convert_office_to_pdf_safe(input_path, output_path, office_app)
            self.window.after(0, self._conversion_done, True, output_path)
        except Exception as e:
            self.window.after(0, self._conversion_done, False, str(e))

    def _conversion_done(self, success, info):
        self._converting = False
        self._set_ui_state(converting=False)

        if success:
            messagebox.showinfo("Success", f"PDF saved as:\n{info}")
        else:
            error_text = f"Conversion failed.\n\n{info}"
            if "excel" in info.lower() or "word" in info.lower() or "powerpoint" in info.lower():
                app_hint = "Microsoft Office"
            else:
                app_hint = "the appropriate Office application"
            error_text += f"\n\nTip: Make sure {app_hint} is installed."
            messagebox.showerror("Error", error_text)

    def _set_ui_state(self, converting):
        if converting:
            self.select_btn.config(state="disabled")
            self.convert_btn.config(state="disabled")
            self.window.config(cursor="watch")
            self.status_var.set("Converting... Please wait")
            self.progress.pack(pady=(5, 0))
            self.progress.start(10)
        else:
            self.select_btn.config(state="normal")
            if self.filepath.get():
                self.convert_btn.config(state="normal")
            else:
                self.convert_btn.config(state="disabled")
            self.window.config(cursor="")
            self.status_var.set("")
            self.progress.stop()
            self.progress.pack_forget()

    def run(self):
        self.window.mainloop()


if __name__ == "__main__":
    app = OfficeToPDFConverter()
    app.run()