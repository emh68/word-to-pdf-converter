import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext, messagebox
import os
from docx2pdf import convert as docx2pdf_convert


class DocxToPdfConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DOCX to PDF Converter")
        self.root.geometry("900x600")
        self.root.configure(bg="#2e2e2e")

        self.files = {}  # {input_path: output_path}
        self.create_widgets()

    # Creates the buttons and labels using tkinter
    def create_widgets(self):
        tk.Label(self.root, text="DOCX to PDF Converter (Multiple Files)",
                 font=("Segoe UI", 16, "bold"), bg="#2e2e2e", fg="white").pack(pady=(10, 0))

        tk.Label(self.root,
                 text="💡 Instructions: Select one or more Word documents (.docx) to convert to PDF.\n"
                      "Use 'Set Output Path' to choose where converted files will be saved. Then click 'Convert All'.",
                 font=("Segoe UI", 10), bg="#2e2e2e", fg="lightgray").pack(pady=(0, 10))

        frame = tk.Frame(self.root, bg="#2e2e2e")
        frame.pack(fill="both", expand=True, padx=10, pady=5)

        btn_frame = tk.Frame(frame, bg="#2e2e2e")
        btn_frame.pack(fill="x")

        tk.Button(btn_frame, text="Add Word Files", command=self.add_files,
                  bg="#007acc", fg="white").pack(side="left", padx=5)
        tk.Button(btn_frame, text="Remove Selected", command=self.remove_selected,
                  bg="#d9534f", fg="white").pack(side="left", padx=5)
        tk.Button(btn_frame, text="Set Output Path for Selected", command=self.set_output_path,
                  bg="#ffc107", fg="black").pack(side="left", padx=5)
        tk.Button(btn_frame, text="Convert All", command=self.convert_all,
                  bg="#28a745", fg="white").pack(side="right", padx=5)

        # Creates window to view input file to be converted and output file location and name
        tree_frame = tk.Frame(frame, bg="#2e2e2e")
        tree_frame.pack(fill="both", expand=True, pady=(10, 0))

        columns = ("Input File", "Output File")
        self.tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings", selectmode="extended")
        self.tree.pack(side="left", fill="both", expand=True)

        # Creates scrollbar for display window
        scrollbar = ttk.Scrollbar(
            tree_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.heading("Input File", text="Input File")
        self.tree.heading("Output File", text="Output File")
        self.tree.column("Input File", anchor="w", width=350)
        self.tree.column("Output File", anchor="w", width=350)

        self.log_text = scrolledtext.ScrolledText(self.root, height=14, bg="#1e1e1e", fg="white",
                                                  insertbackground="white", font=("Consolas", 10))
        self.log_text.pack(fill="both", padx=10, pady=10)

    def log(self, message):
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")

    # Method for adding files (curently only supports .docx file types)
    def add_files(self):
        paths = filedialog.askopenfilenames(
            filetypes=[("Word Documents", "*.docx")])
        for path in paths:
            if path not in self.files:
                default_output = os.path.splitext(path)[0] + ".pdf"
                self.files[path] = default_output
                self.tree.insert("", "end", values=(path, default_output))
                self.log(f"Added: {path}")

    # Method to remove files
    def remove_selected(self):
        selected_items = self.tree.selection()
        if not selected_items:
            self.log("Please select one or more files to remove.")
            return

        for item in selected_items:
            input_path = self.tree.item(item, "values")[0]
            if input_path in self.files:
                del self.files[input_path]
            self.tree.delete(item)
            self.log(f"Removed: {input_path}")

    # Method to set the output file location after conversion to .pdf
    def set_output_path(self):
        selected_items = self.tree.selection()
        if not selected_items:
            self.log("Please select one or more files to set output path.")
            return

        folder = filedialog.askdirectory()
        if not folder:
            return

        for item in selected_items:
            input_path = self.tree.item(item, "values")[0]
            filename = os.path.splitext(
                os.path.basename(input_path))[0] + ".pdf"
            new_output = os.path.join(folder, filename)
            self.files[input_path] = new_output
            self.tree.item(item, values=(input_path, new_output))
            self.log(f"Updated output path for: {input_path} -> {new_output}")

    # Method to convert files from .docx to .pdf using docx2pdf
    def convert_all(self):
        if not self.files:
            self.log("Please add at least one .docx file to convert.")
            return

        # Show warning dialog
        response = messagebox.askokcancel(
            "WARNING",
            "Please SAVE all open Word documents and CLOSE Microsoft Word before continuing!!!\n\n"
            "Click OK to proceed or Cancel to abort."
        )
        if not response:
            self.log("Conversion cancelled by user.")
            return

        self.log("Starting conversion process...")

        for input_path, output_path in self.files.items():
            self.log(f"Converting '{input_path}' to PDF...")
            try:
                docx2pdf_convert(input_path, output_path)
                self.log(f"✅ Successfully converted: {output_path}")
            except Exception as e:
                self.log(f"❌ Failed converting '{input_path}': {e}")

        self.log("Conversion process finished.")


if __name__ == "__main__":
    root = tk.Tk()
    app = DocxToPdfConverterApp(root)
    root.mainloop()
