"""
update_logistics_gui.py
-----------------------
Interface gráfica para o update_logistics.py

Uso:
    python update_logistics_gui.py

Sem dependências extra — tkinter já vem com Python.
"""

import tkinter as tk
from tkinter import filedialog, scrolledtext
import threading
import sys
import os

# Importa a lógica do script principal
# (ambos os ficheiros devem estar na mesma pasta)
import update_logistics as core


# ── Redirecionar print() para a caixa de texto ────────────────────────────────

class TextRedirector:
    """Redireciona stdout para o widget de texto da GUI."""
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        self.widget.configure(state="normal")
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)        # auto-scroll
        self.widget.configure(state="disabled")
        self.widget.update_idletasks() # forçar refresh imediato

    def flush(self):
        pass


# ── GUI ───────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Logistics Management Updater  V0.1")
        self.resizable(False, False)
        self.configure(padx=16, pady=16)

        # Paths selecionados
        self.pdf_path   = tk.StringVar()
        self.excel_path = tk.StringVar(value=core.EXCEL_PATH)

        self._build_ui()

    def _build_ui(self):
        # ── PDF row ───────────────────────────────────────────────────────────
        tk.Label(self, text="DHL Invoice (PDF)", anchor="w").grid(
            row=0, column=0, sticky="w", pady=(0, 4)
        )
        tk.Entry(self, textvariable=self.pdf_path, width=54, state="readonly").grid(
            row=1, column=0, sticky="ew", padx=(0, 8)
        )
        tk.Button(self, text="Browse…", command=self._pick_pdf, width=10).grid(
            row=1, column=1
        )

        # ── Excel row ─────────────────────────────────────────────────────────
        tk.Label(self, text="Logistics Workbook (XLSX)", anchor="w").grid(
            row=2, column=0, sticky="w", pady=(12, 4)
        )
        tk.Entry(self, textvariable=self.excel_path, width=54, state="readonly").grid(
            row=3, column=0, sticky="ew", padx=(0, 8)
        )
        tk.Button(self, text="Browse…", command=self._pick_excel, width=10).grid(
            row=3, column=1
        )

        # ── Run button ────────────────────────────────────────────────────────
        self.run_btn = tk.Button(
            self, text="▶  Run", command=self._run,
            width=14, height=2,
            bg="#2563EB", fg="white", activebackground="#1D4ED8",
            relief="flat", cursor="hand2",
        )
        self.run_btn.grid(row=4, column=0, columnspan=2, pady=(16, 12))

        # ── Output ────────────────────────────────────────────────────────────
        tk.Label(self, text="Output", anchor="w").grid(
            row=5, column=0, sticky="w", pady=(0, 4)
        )
        self.output = scrolledtext.ScrolledText(
            self, width=70, height=18,
            state="disabled", bg="#111827", fg="#F9FAFB",
            font=("Courier New", 10), relief="flat",
        )
        self.output.grid(row=6, column=0, columnspan=2)

    # ── File pickers ──────────────────────────────────────────────────────────

    def _pick_pdf(self):
        path = filedialog.askopenfilename(
            title="Select DHL Invoice PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if path:
            self.pdf_path.set(path)

    def _pick_excel(self):
        path = filedialog.askopenfilename(
            title="Select Logistics Workbook",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.excel_path.set(path)

    # ── Run logic ─────────────────────────────────────────────────────────────

    def _run(self):
        pdf   = self.pdf_path.get().strip()
        excel = self.excel_path.get().strip()

        if not pdf:
            self._log("ERROR: Please select a PDF file.\n")
            return
        if not excel:
            self._log("ERROR: Please select the Excel workbook.\n")
            return
        if not os.path.exists(pdf):
            self._log(f"ERROR: PDF not found:\n  {pdf}\n")
            return
        if not os.path.exists(excel):
            self._log(f"ERROR: Workbook not found:\n  {excel}\n")
            return

        # Clear output
        self.output.configure(state="normal")
        self.output.delete("1.0", tk.END)
        self.output.configure(state="disabled")

        # Disable button while running
        self.run_btn.configure(state="disabled", text="Running…")

        # Run in background thread so the UI doesn't freeze
        threading.Thread(
            target=self._worker, args=(pdf, excel), daemon=True
        ).start()

    def _worker(self, pdf_path, excel_path):
        # Patch EXCEL_PATH in the core module so it uses the GUI selection
        core.EXCEL_PATH = excel_path

        # Redirect stdout to the output box
        original_stdout = sys.stdout
        sys.stdout = TextRedirector(self.output)

        try:
            # Simulate sys.argv so main() finds the PDF path
            sys.argv = ["update_logistics.py", pdf_path]
            core.main()
        except SystemExit:
            pass
        except Exception as e:
            print(f"\nUnexpected error: {e}")
        finally:
            sys.stdout = original_stdout
            # Re-enable button on the main thread
            self.after(0, lambda: self.run_btn.configure(
                state="normal", text="▶  Run"
            ))

    def _log(self, text):
        self.output.configure(state="normal")
        self.output.insert(tk.END, text)
        self.output.configure(state="disabled")


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    App().mainloop()