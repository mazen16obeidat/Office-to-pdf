import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD  # for drag & drop

try:
    import win32com.client
    import win32com.client.dynamic
    import pythoncom
except ImportError:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Missing Library",
        "pywin32 is required.\nRun: pip install pywin32"
    )
    sys.exit(1)

# ---------- Configuration ----------
SUPPORTED_FORMATS = {
    # Word
    ".docx": "Word", ".doc": "Word", ".docm": "Word",
    ".dotx": "Word", ".dotm": "Word", ".rtf": "Word", ".odt": "Word",
    # Excel
    ".xlsx": "Excel", ".xls": "Excel", ".xlsm": "Excel",
    ".xltx": "Excel", ".xltm": "Excel", ".csv": "Excel",
    # PowerPoint
    ".pptx": "PowerPoint", ".ppt": "PowerPoint", ".pptm": "PowerPoint",
    ".potx": "PowerPoint", ".potm": "PowerPoint",
}

FILE_DIALOG_PATTERNS = [
    ("All Supported Files",
     "*.docx;*.doc;*.docm;*.dotx;*.dotm;*.rtf;*.odt;"
     "*.xlsx;*.xls;*.xlsm;*.xltx;*.xltm;*.csv;"
     "*.pptx;*.ppt;*.pptm;*.potx;*.potm"),
    ("Word Documents", "*.docx;*.doc;*.docm;*.dotx;*.dotm;*.rtf;*.odt"),
    ("Excel Spreadsheets", "*.xlsx;*.xls;*.xlsm;*.xltx;*.xltm;*.csv"),
    ("PowerPoint Presentations", "*.pptx;*.ppt;*.pptm;*.potx;*.potm"),
    ("All Files", "*.*"),
]

# ---------- Core Conversion Logic ----------
def convert_office_to_pdf_safe(input_path, output_path, office_app):
    """Convert a file using the appropriate Office application in a thread-safe context."""
    abs_input = os.path.abspath(input_path)
    abs_output = os.path.abspath(output_path)

    pythoncom.CoInitialize()

    app = None
    doc = None
    try:
        app = win32com.client.dynamic.Dispatch(f"{office_app}.Application")

        if office_app == "Word":
            app.Visible = False
            app.DisplayAlerts = 0  # wdAlertsNone
            # Positional args: FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
            doc = app.Documents.Open(abs_input, False, True, False)
            doc.ExportAsFixedFormat(abs_output, 17)  # wdExportFormatPDF = 17

        elif office_app == "Excel":
            app.Visible = False
            app.DisplayAlerts = False
            # Positional args: Filename, UpdateLinks, ReadOnly
            doc = app.Workbooks.Open(abs_input, 0, True)
            doc.ExportAsFixedFormat(0, abs_output)  # xlTypePDF = 0

        elif office_app == "PowerPoint":
            # PowerPoint REQUIRES app.Visible = 1 for the render engine to process slides
            app.Visible = 1
            
            # Open args: FileName, ReadOnly (-1), Untitled (0), WithWindow (-1)
            doc = app.Presentations.Open(abs_input, -1, 0, -1)
            
            # SaveAs with format 32 (ppSaveAsPDF) is universally supported
            doc.SaveAs(abs_output, 32)

        return True

    finally:
        if doc:
            try:
                doc.Close()
            except Exception:
                pass
        if app:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

# ---------- GUI Application ----------
class OfficeToPDFConverter:
    def __init__(self):
        self.window = TkinterDnD.Tk()
        self.window.title("Family PDF Tool v2.0")
        self.window.geometry("600x360")
        self.window.resizable(False, False)

        try:
            self.window.iconbitmap(default="C:\\Windows\\System32\\shell32.dll")
        except Exception:
            pass

        for family in ("Segoe UI", "Tahoma", "Arial"):
            try:
                self.window.option_add("*Font", (family, 10))
                break
            except Exception:
                continue

        self.filepath = tk.StringVar()
        self.filename_display = tk.StringVar(value="Drop a file here or click 'Select File'")
        self.filetype_display = tk.StringVar()

        self._converting = False
        self.open_after_var = tk.BooleanVar(value=True)

        self._setup_ui()
        self._setup_drag_drop()

        self.window.bind("<Control-t>", lambda e: self._diagnose_office())

    # ---------- UI Construction ----------
    def _setup_ui(self):
        main_frame = tk.Frame(self.window, padx=25, pady=20)
        main_frame.pack(expand=True, fill="both")

        title = tk.Label(main_frame, text="Office to PDF Converter",
                         font=("Segoe UI", 15, "bold"))
        title.pack(pady=(0, 5))

        subtitle = tk.Label(main_frame, text="Word • Excel • PowerPoint",
                            font=("Segoe UI", 9), fg="gray")
        subtitle.pack(pady=(0, 15))

        self.file_label = tk.Label(main_frame, textvariable=self.filename_display,
                                   wraplength=500, anchor="center", justify="center",
                                   bg="#f0f0f0", relief="ridge", padx=10, pady=8)
        self.file_label.pack(pady=(0, 5), fill="x")

        self.type_label = tk.Label(main_frame, textvariable=self.filetype_display,
                                   fg="#0078D4", font=("Segoe UI", 9, "bold"))
        self.type_label.pack(pady=(0, 5))

        self.status_var = tk.StringVar()
        self.status_label = tk.Label(main_frame, textvariable=self.status_var, fg="gray")
        self.status_label.pack(pady=(0, 3))

        self.progress = ttk.Progressbar(main_frame, mode="indeterminate", length=300)

        self.open_cb = tk.Checkbutton(main_frame, text="Open PDF after conversion",
                                      variable=self.open_after_var)
        self.open_cb.pack(pady=(5, 0))

        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(pady=15)

        self.select_btn = tk.Button(btn_frame, text="Select File", width=14,
                                    command=self.select_file)
        self.select_btn.grid(row=0, column=0, padx=4)

        self.clear_btn = tk.Button(btn_frame, text="Clear", width=8,
                                   command=self.clear_file, state="disabled")
        self.clear_btn.grid(row=0, column=1, padx=4)

        self.convert_btn = tk.Button(btn_frame, text="Convert to PDF", width=16,
                                     command=self.start_conversion, state="disabled",
                                     bg="#0078D4", fg="white", font=("Segoe UI", 10, "bold"))
        self.convert_btn.grid(row=0, column=2, padx=4)

    def _setup_drag_drop(self):
        self.window.drop_target_register(DND_FILES)
        self.window.dnd_bind("<<Drop>>", self.on_drop)

    def on_drop(self, event):
        files = self.window.tk.splitlist(event.data)
        if files:
            self._set_file(files[0])

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Select an Office Document",
            filetypes=FILE_DIALOG_PATTERNS
        )
        if path:
            self._set_file(path)

    def _set_file(self, path):
        ext = os.path.splitext(path)[1].lower()
        office_app = SUPPORTED_FORMATS.get(ext)

        if office_app:
            self.filepath.set(path)
            self.filename_display.set(os.path.basename(path))
            self.filetype_display.set(f"Detected: {office_app} document")
            self.convert_btn.config(state="normal")
            self.clear_btn.config(state="normal")
        else:
            messagebox.showwarning(
                "Unsupported Format",
                f"The file type '{ext}' is not supported.\n\n"
                "Supported formats: Word, Excel, PowerPoint.\n"
                "See the list in the file dialog."
            )

    def clear_file(self):
        self.filepath.set("")
        self.filename_display.set("Drop a file here or click 'Select File'")
        self.filetype_display.set("")
        self.convert_btn.config(state="disabled")
        self.clear_btn.config(state="disabled")

    # ---------- Conversion Flow ----------
    def start_conversion(self):
        if self._converting:
            return

        input_path = self.filepath.get()
        if not input_path:
            return

        output_path = os.path.splitext(input_path)[0] + ".pdf"

        if os.path.exists(output_path):
            if not messagebox.askyesno(
                "File Exists",
                f"'{os.path.basename(output_path)}' already exists.\nOverwrite?"
            ):
                return

        ext = os.path.splitext(input_path)[1].lower()
        office_app = SUPPORTED_FORMATS.get(ext)

        if not office_app:
            messagebox.showerror("Error", "Unsupported file type.")
            return

        self._converting = True
        self._set_ui_state(converting=True)

        thread = threading.Thread(
            target=self._convert_worker,
            args=(input_path, output_path, office_app),
            daemon=True
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
            msg = f"PDF saved as:\n{info}"
            if self.open_after_var.get():
                os.startfile(info)
            messagebox.showinfo("Success", msg)
        else:
            error_text = f"Conversion failed.\n\n{info}"
            lower_info = info.lower()
            if "excel" in lower_info or "xls" in lower_info:
                hint = "Microsoft Excel"
            elif "powerpoint" in lower_info or "ppt" in lower_info:
                hint = "Microsoft PowerPoint"
            elif "word" in lower_info or "doc" in lower_info:
                hint = "Microsoft Word"
            else:
                hint = "the appropriate Office application"

            error_text += f"\n\n💡 Make sure {hint} is installed and the file is accessible."
            error_text += "\nPress Ctrl+T to check which Office apps are available."
            messagebox.showerror("Error", error_text)

    # ---------- UI State Management ----------
    def _set_ui_state(self, converting):
        if converting:
            self.select_btn.config(state="disabled")
            self.clear_btn.config(state="disabled")
            self.convert_btn.config(state="disabled")
            self.window.config(cursor="watch")
            self.status_var.set("Converting... Please wait")
            self.progress.pack(pady=(5, 0))
            self.progress.start(10)
        else:
            self.select_btn.config(state="normal")
            if self.filepath.get():
                self.clear_btn.config(state="normal")
            else:
                self.clear_btn.config(state="disabled")
            self.convert_btn.config(state="normal")
            self.window.config(cursor="")
            self.status_var.set("")
            self.progress.stop()
            self.progress.pack_forget()

    # ---------- Hidden Diagnostic (Ctrl+T) ----------
    def _diagnose_office(self):
        pythoncom.CoInitialize()
        results = []
        for app_name in ["Word", "Excel", "PowerPoint"]:
            try:
                app = win32com.client.dynamic.Dispatch(f"{app_name}.Application")
                ver = app.Version
                results.append(f"✅ {app_name}: Found (version {ver})")
                app.Quit()
            except Exception:
                results.append(f"❌ {app_name}: Not found")
        pythoncom.CoUninitialize()
        messagebox.showinfo("Office Availability", "\n".join(results))

    def run(self):
        self.window.mainloop()

# ---------- Entry Point ----------
if __name__ == "__main__":
    app = OfficeToPDFConverter()
    app.run()
