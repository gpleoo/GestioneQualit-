"""
Gestione Qualita - Auto-Fill Excel da PDF DOP

Applicazione desktop per compilare automaticamente il file Excel
"Controllo Prefabbricazione" estraendo i dati da uno o piu PDF DOP
(Dichiarazione di Prestazione CE).

Supporta il caricamento di piu PDF contemporaneamente con generazione
di schede Excel numerate progressivamente (ordinate per data DOP).

Uso:
    python app.py
"""

import os
import sys
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from pdf_parser import extract_dop_data
from excel_filler import fill_excel, get_cell_map


def _parse_date(date_str: str) -> datetime:
    """Converte una stringa data DD/MM/YYYY in datetime per ordinamento."""
    try:
        return datetime.strptime(date_str, "%d/%m/%Y")
    except (ValueError, TypeError):
        return datetime.max  # date non parsabili vanno in fondo


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gestione Qualita - Compilazione Excel da DOP")
        self.geometry("800x680")
        self.resizable(True, True)
        self.configure(bg="#f0f0f0")

        self.pdf_paths = []  # lista di percorsi PDF
        self.excel_path = tk.StringVar()
        self.dop_data_list = []  # lista di (pdf_path, dop_data) ordinata per data

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

        # PDF (multipli)
        ttk.Label(files_frame, text="PDF DOP:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.pdf_label = ttk.Label(files_frame, text="Nessun PDF selezionato", width=55, anchor=tk.W)
        self.pdf_label.grid(row=0, column=1, padx=5, pady=5)
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

        # Pulsanti azione
        btn_frame = tk.Frame(self, bg="#f0f0f0")
        btn_frame.pack(fill=tk.X, padx=15, pady=10)
        ttk.Button(
            btn_frame,
            text="1. Estrai dati dai PDF",
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
            height=20,
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
        self.status_var = tk.StringVar(value="Pronto. Seleziona uno o piu PDF DOP e un file Excel template.")
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
        paths = filedialog.askopenfilenames(
            title="Seleziona uno o piu PDF DOP",
            filetypes=[("File PDF", "*.pdf"), ("Tutti i file", "*.*")],
        )
        if paths:
            self.pdf_paths = list(paths)
            n = len(self.pdf_paths)
            if n == 1:
                self.pdf_label.configure(text=os.path.basename(self.pdf_paths[0]))
            else:
                self.pdf_label.configure(text=f"{n} PDF selezionati")
            self.status_var.set(f"{n} PDF selezionato/i.")

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Seleziona il file Excel template",
            filetypes=[("File Excel", "*.xlsx *.xls"), ("Tutti i file", "*.*")],
        )
        if path:
            self.excel_path.set(path)
            self.status_var.set(f"Excel selezionato: {os.path.basename(path)}")

    def _extract_data(self):
        if not self.pdf_paths:
            messagebox.showwarning("Attenzione", "Seleziona prima uno o piu file PDF DOP.")
            return

        # Verifica che tutti i file esistano
        for pdf in self.pdf_paths:
            if not os.path.isfile(pdf):
                messagebox.showerror("Errore", f"File non trovato:\n{pdf}")
                return

        try:
            self.status_var.set("Estrazione dati in corso...")
            self.update_idletasks()

            # Estrai dati da ogni PDF
            extracted = []
            errors = []
            for pdf in self.pdf_paths:
                try:
                    data = extract_dop_data(pdf)
                    extracted.append((pdf, data))
                except Exception as e:
                    errors.append(f"{os.path.basename(pdf)}: {e}")

            if errors:
                messagebox.showwarning(
                    "Attenzione",
                    f"Errori su {len(errors)} PDF:\n\n" + "\n".join(errors)
                )

            if not extracted:
                messagebox.showerror("Errore", "Nessun PDF elaborato con successo.")
                return

            # Ordina per data DDT (la meno recente prima = scheda n.1)
            extracted.sort(key=lambda x: _parse_date(x[1].get("data_ddt", "")))

            self.dop_data_list = extracted
            self._show_preview()
            self.status_var.set(
                f"{len(extracted)} PDF estratti con successo! "
                f"Ordinati per data (dal meno recente). Premi 'Compila Excel'."
            )
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
            "posizioni_stringa": "Posizioni (-> cella D2)",
        }

        for idx, (pdf_path, data) in enumerate(self.dop_data_list, start=1):
            self.preview_text.insert(
                tk.END,
                f"{'=' * 50}\n"
                f"  SCHEDA N. {idx:03d} - {os.path.basename(pdf_path)}\n"
                f"{'=' * 50}\n\n"
            )

            for key, label in labels.items():
                value = data.get(key, "")
                self.preview_text.insert(tk.END, f"  {label}:\n")
                self.preview_text.insert(tk.END, f"    {value if value else '(non trovato)'}\n\n")

            # Mostra mappatura Excel
            self.preview_text.insert(tk.END, "  --- Mappatura Excel ---\n")
            for label, field in get_cell_map().items():
                value = data.get(field, "(non trovato)")
                self.preview_text.insert(tk.END, f"    {label} -> {value}\n")
            self.preview_text.insert(tk.END, "\n\n")

        self.preview_text.insert(
            tk.END,
            f"Totale: {len(self.dop_data_list)} schede da generare\n"
            f"Ordine: per data DOP crescente (la meno recente = Scheda 001)\n"
        )

        self.preview_text.configure(state=tk.DISABLED)

    def _fill_and_save(self):
        if not self.dop_data_list:
            messagebox.showwarning("Attenzione", "Prima estrai i dati dai PDF (pulsante 1).")
            return

        excel = self.excel_path.get()
        if not excel:
            messagebox.showwarning("Attenzione", "Seleziona un file Excel template.")
            return

        if not os.path.isfile(excel):
            messagebox.showerror("Errore", f"File non trovato:\n{excel}")
            return

        # Scegli cartella di output
        output_dir = filedialog.askdirectory(
            title="Seleziona la cartella dove salvare i file Excel compilati",
            initialdir=os.path.dirname(excel),
        )
        if not output_dir:
            return

        try:
            self.status_var.set("Compilazione Excel in corso...")
            self.update_idletasks()

            base_name = os.path.splitext(os.path.basename(excel))[0]
            generated = []

            for idx, (pdf_path, data) in enumerate(self.dop_data_list, start=1):
                output_name = f"{base_name}_{idx:03d}.xlsx"
                output_path = os.path.join(output_dir, output_name)
                fill_excel(excel, output_path, data)
                generated.append(output_name)

            self.status_var.set(f"{len(generated)} file Excel generati in {output_dir}")

            file_list = "\n".join(f"  {name}" for name in generated)
            messagebox.showinfo(
                "Completato",
                f"{len(generated)} schede Excel generate con successo!\n\n"
                f"Cartella: {output_dir}\n\n"
                f"File generati:\n{file_list}",
            )
        except Exception as e:
            messagebox.showerror("Errore compilazione", f"Errore durante la compilazione:\n{e}")
            self.status_var.set("Errore durante la compilazione Excel.")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
