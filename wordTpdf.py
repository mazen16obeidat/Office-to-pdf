import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# We use win32com directly instead of docx2pdf,
# so we can open the document as read-only.
try:
    import win32com.client
except ImportError:
    messagebox.showerror(
        "Missing Library",
        "pywin32 is required.\nRun: pip install pywin32"
    )
    raise SystemExit


def convert_docx_to_pdf_safe(docx_path, pdf_path):
    """
    Convert a .docx to PDF using Word in read-only mode.
    This works even if the document is already open in Word.
    """
    word = None
    doc = None
    try:
        # Start Word invisibly
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        # Open the document read-only – avoids locking conflicts
        doc = word.Documents.Open(docx_path, ReadOnly=True)

        # Export as PDF (format type 17 = PDF)
        doc.ExportAsFixedFormat(pdf_path, ExportFormat=17)

        # Success
        return True

    except Exception as e:
        # Re-raise so the GUI can show it
        raise e

    finally:
        # Clean up COM objects in reverse order
        if doc:
            doc.Close(SaveChanges=False)
        if word:
            word.Quit()


class WordToPDFConverter:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Family PDF Tool")
        self.window.geometry("520x280")
        self.window.resizable(False, False)

        default_font = ("Segoe UI", 10)
        self.window.option_add("*Font", default_font)

        self.filepath = tk.StringVar()
        self.filename_display = tk.StringVar(value="No file selected")

        self._converting = False

        self._setup_ui()

    def _setup_ui(self):
        main_frame = tk.Frame(self.window, padx=30, pady=20)
        main_frame.pack(expand=True, fill="both")

        title = tk.Label(
            main_frame,
            text="Word to PDF Converter",
            font=("Segoe UI", 14, "bold"),
        )
        title.pack(pady=(0, 15))

        self.file_label = tk.Label(
            main_frame,
            textvariable=self.filename_display,
            wraplength=400,
            anchor="center",
            justify="center",
        )
        self.file_label.pack(pady=(0, 5))

        self.status_var = tk.StringVar()
        self.status_label = tk.Label(
            main_frame, textvariable=self.status_var, fg="gray"
        )
        self.status_label.pack(pady=(0, 5))

        self.progress = ttk.Progressbar(
            main_frame, mode="indeterminate", length=200
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
        path = filedialog.askopenfilename(
            title="Select a Word Document",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        if path:
            self.filepath.set(path)
            self.filename_display.set(os.path.basename(path))
            self.convert_btn.config(state="normal")
        else:
            self.filepath.set("")
            self.filename_display.set("No file selected")
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
                f"{output_path}\nalready exists.\nOverwrite?",
            )
            if not overwrite:
                return

        self._converting = True
        self._set_ui_state(converting=True)

        thread = threading.Thread(
            target=self._convert_worker, args=(input_path, output_path), daemon=True
        )
        thread.start()

    def _convert_worker(self, docx_path, pdf_path):
        try:
            convert_docx_to_pdf_safe(docx_path, pdf_path)
            self.window.after(0, self._conversion_done, True, pdf_path)
        except Exception as e:
            self.window.after(0, self._conversion_done, False, str(e))

    def _conversion_done(self, success, info):
        self._converting = False
        self._set_ui_state(converting=False)

        if success:
            messagebox.showinfo("Success", f"PDF saved as:\n{info}")
        else:
            error_text = f"Conversion failed.\n\n{info}"
            if "word" in info.lower() or "com" in info.lower():
                error_text += "\n\nTip: Microsoft Word must be installed."
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
    app = WordToPDFConverter()
    app.run()