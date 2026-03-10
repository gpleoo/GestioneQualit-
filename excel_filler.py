"""
Modulo per la compilazione automatica del file Excel
"Controllo Prefabbricazione" a partire dai dati estratti dal PDF DOP.

Mappatura celle Excel (basata sul template):
- D2: Elemento assiemato (codici posizione)
- G10: Data controllo (riga 5 - Esame visivo saldature)
- I14: Data compilazione
"""

from copy import copy
from openpyxl import load_workbook


# Mappatura celle: chiave = nome campo, valore = cella Excel
CELL_MAP = {
    "elemento_assiemato": "D2",   # Codici posizione prodotto
    "data_controllo": "G10",      # Data nella riga "Esame visivo" (riga 5)
    "data_compilazione": "I14",   # Data compilazione in fondo
}


def fill_excel(template_path: str, output_path: str, dop_data: dict) -> str:
    """
    Compila il file Excel template con i dati estratti dal PDF DOP.

    Args:
        template_path: Percorso del file Excel template.
        output_path: Percorso dove salvare il file compilato.
        dop_data: Dizionario con i dati estratti dal PDF (output di extract_dop_data).

    Returns:
        Percorso del file salvato.
    """
    wb = load_workbook(template_path)
    ws = wb.active  # Primo foglio

    # Compila Elemento assiemato (D2) con i codici posizione
    if dop_data.get("posizioni_stringa"):
        _write_cell(ws, CELL_MAP["elemento_assiemato"], dop_data["posizioni_stringa"])

    # Compila Data controllo (G10) con la data del DDT
    if dop_data.get("data_ddt"):
        _write_cell(ws, CELL_MAP["data_controllo"], dop_data["data_ddt"])

    # Compila Data compilazione (I14) con la data del DDT
    if dop_data.get("data_ddt"):
        _write_cell(ws, CELL_MAP["data_compilazione"], dop_data["data_ddt"])

    wb.save(output_path)
    return output_path


def _write_cell(ws, cell_ref: str, value):
    """Scrive un valore in una cella preservando la formattazione esistente."""
    cell = ws[cell_ref]
    cell.value = value


def get_cell_map() -> dict:
    """Restituisce la mappatura celle corrente per visualizzazione nella GUI."""
    return {
        "Elemento assiemato (D2)": "posizioni_stringa",
        "Data controllo (G10)": "data_ddt",
        "Data compilazione (I14)": "data_ddt",
    }


if __name__ == "__main__":
    print("=== Mappatura celle Excel ===")
    for label, field in get_cell_map().items():
        print(f"  {label} ← campo PDF: {field}")
