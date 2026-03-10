"""
Gestione Qualità - Auto-Fill Excel da PDF DOP

Applicazione desktop per compilare automaticamente il file Excel
"Controllo Prefabbricazione" estraendo i dati da un PDF DOP
(Dichiarazione di Prestazione CE).

Uso:
    python app.py
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pdf_parser import extract_dop_data
from excel_filler import fill_excel, get_cell_map


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gestione Qualità - Compilazione Excel da DOP")
        self.geometry("750x620")
        self.resizable(True, True)
        self.configure(bg="#f0f0f0")

        self.pdf_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.dop_data = {}

        self._build_ui()

    def _build_ui(self):
        # Titolo
        title_frame = tk.Frame(self, bg="#2c3e50", pady=10)
        title_frame.pack(fill=tk.X)
        tk.Label(
            title_frame,
            text="Compilazione Automatica Excel da PDF DOP",
            font=("Segoe UI", 14, "bold"),
            fg="white",
            bg="#2c3e50",
        ).pack()
        tk.Label(
            title_frame,
            text="CM SRL - Controllo Prefabbricazione",
            font=("Segoe UI", 10),
            fg="#bdc3c7",
            bg="#2c3e50",
        ).pack()

        # Frame selezione file
        files_frame = ttk.LabelFrame(self, text="  Selezione File  ", padding=15)
        files_frame.pack(fill=tk.X, padx=15, pady=(15, 5))

        # PDF
        ttk.Label(files_frame, text="PDF DOP:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(files_frame, textvariable=self.pdf_path, width=55).grid(
            row=0, column=1, padx=5, pady=5
        )
        ttk.Button(files_frame, text="Sfoglia...", command=self._browse_pdf).grid(
            row=0, column=2, pady=5
        )

        # Excel template
        ttk.Label(files_frame, text="Excel Template:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(files_frame, textvariable=self.excel_path, width=55).grid(
            row=1, column=1, padx=5, pady=5
        )
        ttk.Button(files_frame, text="Sfoglia...", command=self._browse_excel).grid(
            row=1, column=2, pady=5
        )

        # Pulsante Estrai
        btn_frame = tk.Frame(self, bg="#f0f0f0")
        btn_frame.pack(fill=tk.X, padx=15, pady=10)
        ttk.Button(
            btn_frame,
            text="1. Estrai dati dal PDF",
            command=self._extract_data,
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            btn_frame,
            text="2. Compila Excel e Salva",
            command=self._fill_and_save,
        ).pack(side=tk.LEFT, padx=5)

        # Frame anteprima dati
        preview_frame = ttk.LabelFrame(self, text="  Anteprima Dati Estratti  ", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(5, 15))

        self.preview_text = tk.Text(
            preview_frame,
            height=18,
            font=("Consolas", 10),
            wrap=tk.WORD,
            state=tk.DISABLED,
            bg="#fafafa",
        )
        scrollbar = ttk.Scrollbar(preview_frame, command=self.preview_text.yview)
        self.preview_text.configure(yscrollcommand=scrollbar.set)
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Barra di stato
        self.status_var = tk.StringVar(value="Pronto. Seleziona un PDF DOP e un file Excel template.")
        status_bar = tk.Label(
            self,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W,
            bg="#ecf0f1",
            font=("Segoe UI", 9),
            padx=10,
        )
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _browse_pdf(self):
        path = filedialog.askopenfilename(
            title="Seleziona il PDF DOP",
            filetypes=[("File PDF", "*.pdf"), ("Tutti i file", "*.*")],
        )
        if path:
            self.pdf_path.set(path)
            self.status_var.set(f"PDF selezionato: {os.path.basename(path)}")

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Seleziona il file Excel template",
            filetypes=[("File Excel", "*.xlsx *.xls"), ("Tutti i file", "*.*")],
        )
        if path:
            self.excel_path.set(path)
            self.status_var.set(f"Excel selezionato: {os.path.basename(path)}")

    def _extract_data(self):
        pdf = self.pdf_path.get()
        if not pdf:
            messagebox.showwarning("Attenzione", "Seleziona prima un file PDF DOP.")
            return

        if not os.path.isfile(pdf):
            messagebox.showerror("Errore", f"File non trovato:\n{pdf}")
            return

        try:
            self.status_var.set("Estrazione dati in corso...")
            self.update_idletasks()
            self.dop_data = extract_dop_data(pdf)
            self._show_preview()
            self.status_var.set("Dati estratti con successo! Verifica e poi premi 'Compila Excel'.")
        except Exception as e:
            messagebox.showerror("Errore estrazione", f"Errore durante l'estrazione:\n{e}")
            self.status_var.set("Errore durante l'estrazione.")

    def _show_preview(self):
        self.preview_text.configure(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)

        labels = {
            "numero_dop": "Numero DOP",
            "nr_riferimento": "Nr. Riferimento",
            "numero_ddt": "Numero DDT",
            "data_ddt": "Data DDT",
            "tipo_prodotto": "Tipo Prodotto",
            "posizioni_stringa": "Posizioni (→ cella D2)",
            "fabbricante": "Fabbricante",
            "norma": "Norma",
        }

        self.preview_text.insert(tk.END, "═══ DATI ESTRATTI DAL PDF DOP ═══\n\n")
        for key, label in labels.items():
            value = self.dop_data.get(key, "")
            self.preview_text.insert(tk.END, f"  {label}:\n")
            self.preview_text.insert(tk.END, f"    {value if value else '(non trovato)'}\n\n")

        self.preview_text.insert(tk.END, "\n═══ MAPPATURA → EXCEL ═══\n\n")
        for label, field in get_cell_map().items():
            value = self.dop_data.get(field, "(non trovato)")
            self.preview_text.insert(tk.END, f"  {label}\n")
            self.preview_text.insert(tk.END, f"    → {value}\n\n")

        self.preview_text.configure(state=tk.DISABLED)

    def _fill_and_save(self):
        if not self.dop_data:
            messagebox.showwarning("Attenzione", "Prima estrai i dati dal PDF (pulsante 1).")
            return

        excel = self.excel_path.get()
        if not excel:
            messagebox.showwarning("Attenzione", "Seleziona un file Excel template.")
            return

        if not os.path.isfile(excel):
            messagebox.showerror("Errore", f"File non trovato:\n{excel}")
            return

        # Proponi nome file output
        dir_name = os.path.dirname(excel)
        base_name = os.path.splitext(os.path.basename(excel))[0]
        dop_num = self.dop_data.get("numero_dop", "")
        suggested_name = f"{base_name}_compilato_{dop_num}.xlsx" if dop_num else f"{base_name}_compilato.xlsx"

        output = filedialog.asksaveasfilename(
            title="Salva il file Excel compilato",
            initialdir=dir_name,
            initialfile=suggested_name,
            defaultextension=".xlsx",
            filetypes=[("File Excel", "*.xlsx")],
        )
        if not output:
            return

        try:
            self.status_var.set("Compilazione Excel in corso...")
            self.update_idletasks()
            fill_excel(excel, output, self.dop_data)
            self.status_var.set(f"File salvato: {os.path.basename(output)}")
            messagebox.showinfo(
                "Completato",
                f"File Excel compilato e salvato con successo!\n\n{output}",
            )
        except Exception as e:
            messagebox.showerror("Errore compilazione", f"Errore durante la compilazione:\n{e}")
            self.status_var.set("Errore durante la compilazione Excel.")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
