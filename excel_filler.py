"""
Modulo per la compilazione automatica del file Excel
"Controllo Prefabbricazione" a partire dai dati estratti dal PDF DOP.

Utilizza la modifica diretta dell'XML interno al file xlsx (ZIP)
per preservare immagini, formule e formattazione del template originale.

Mappatura celle Excel (basata sul template):
- G2:  Cliente
- D2:  Elemento assiemato (codici posizione)
- D3:  N° commessa
- D4:  Progetto
- G6:  Data collaudo (più recente dal file Excel marcature)
- G10: Data controllo (riga 5 - Esame visivo saldature)
- I14: Data compilazione
"""

import io
import os
import re
import shutil
import zipfile
from datetime import datetime
from xml.etree import ElementTree as ET

import openpyxl


# Mappatura celle: chiave = nome campo, valore = cella Excel
CELL_MAP = {
    "numero_scheda":      "C1",
    "cliente":            "G2",
    "elemento_assiemato": "D2",
    "numero_commessa":    "D3",
    "progetto":           "D4",
    "data_collaudo":      "G6",
    "data_controllo":     "G10",
    "data_compilazione":  "I14",
}

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def fill_excel(
    template_path: str,
    output_path: str,
    dop_data: dict,
    manual_data: dict = None,
    marcature_excel_path: str = "",
    numero_scheda: int = 0,
) -> str:
    """
    Compila il file Excel template con i dati estratti dal PDF DOP.

    Copia il template e modifica solo le celle necessarie a livello XML,
    preservando immagini, formule e formattazione originali.

    Args:
        template_path: Percorso del file Excel template.
        output_path: Percorso dove salvare il file compilato.
        dop_data: Dizionario con i dati estratti dal PDF.
        manual_data: Dizionario con i campi inseriti manualmente (cliente, numero_commessa, progetto).
        marcature_excel_path: Percorso del file Excel con colonna A=marcature, D=date.

    Returns:
        Percorso del file salvato.
    """
    cells_to_write = {}

    if numero_scheda > 0:
        cells_to_write[CELL_MAP["numero_scheda"]] = f"{numero_scheda:03d}"

    if dop_data.get("posizioni_stringa"):
        cells_to_write[CELL_MAP["elemento_assiemato"]] = dop_data["posizioni_stringa"]
    if dop_data.get("data_ddt"):
        cells_to_write[CELL_MAP["data_controllo"]] = dop_data["data_ddt"]
        cells_to_write[CELL_MAP["data_compilazione"]] = dop_data["data_ddt"]

    # Dati inseriti manualmente dall'utente
    if manual_data:
        for field in ("cliente", "numero_commessa", "progetto"):
            val = manual_data.get(field, "").strip()
            if val:
                cells_to_write[CELL_MAP[field]] = val

    # Data più recente dal file Excel marcature → G6
    if marcature_excel_path and dop_data.get("posizioni"):
        data_collaudo = _get_most_recent_date_from_excel(
            marcature_excel_path, dop_data["posizioni"]
        )
        if data_collaudo:
            cells_to_write[CELL_MAP["data_collaudo"]] = data_collaudo

    # Copia il template preservando tutto (immagini, macro, ecc.)
    shutil.copy2(template_path, output_path)

    if cells_to_write:
        _patch_xlsx(output_path, cells_to_write)

    return output_path


def _parse_date_from_string(value: str) -> datetime | None:
    """Estrae la data da stringhe come '1 del 07/04/25' o '2 del 09/04/2025'."""
    match = re.search(r"(\d{2}/\d{2}/\d{2,4})", str(value))
    if not match:
        return None
    date_str = match.group(1)
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None


def _get_most_recent_date_from_excel(excel_path: str, posizioni: list) -> str:
    """
    Legge il file Excel delle marcature e restituisce la data più recente
    tra quelle associate alle posizioni indicate.

    Colonna A: marcatura (es. T2, T4)
    Colonna D: stringa data (es. '1 del 07/04/25')
    """
    if not excel_path or not posizioni or not os.path.isfile(excel_path):
        return ""

    posizioni_set = {str(p).strip().upper() for p in posizioni}

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb.active
    except Exception:
        return ""

    dates = []
    for row in ws.iter_rows(values_only=True):
        col_a = str(row[0]).strip().upper() if row[0] is not None else ""
        col_d = row[3] if len(row) > 3 else None

        if col_a in posizioni_set and col_d is not None:
            # col_d può essere già un datetime (se la cella Excel è di tipo data)
            if isinstance(col_d, datetime):
                dates.append(col_d)
            else:
                parsed = _parse_date_from_string(str(col_d))
                if parsed:
                    dates.append(parsed)

    if not dates:
        return ""
    return max(dates).strftime("%d/%m/%Y")


def _col_to_num(col_str):
    """Converte lettera colonna in numero (A=1, B=2, ..., AA=27)."""
    result = 0
    for c in col_str.upper():
        result = result * 26 + (ord(c) - 64)
    return result


def _parse_cell_ref(ref):
    """Splitta riferimento cella in (colonna_str, riga_int)."""
    m = re.match(r"^([A-Z]+)(\d+)$", ref)
    return m.group(1), int(m.group(2))


def _resolve_merge(ref, merge_ranges):
    """Se ref e' dentro un range unito, restituisce la cella in alto a sinistra."""
    col_s, row = _parse_cell_ref(ref)
    col = _col_to_num(col_s)
    for mr in merge_ranges:
        tl, br = mr.split(":")
        tl_c, tl_r = _parse_cell_ref(tl)
        br_c, br_r = _parse_cell_ref(br)
        if tl_r <= row <= br_r and _col_to_num(tl_c) <= col <= _col_to_num(br_c):
            return tl
    return ref


def _find_sheet_path(zf):
    """Trova il percorso del primo foglio di lavoro nel file xlsx."""
    for name in sorted(zf.namelist()):
        if re.match(r"xl/worksheets/sheet\d+\.xml$", name):
            return name
    return "xl/worksheets/sheet1.xml"


def _patch_xlsx(xlsx_path, cells_to_write):
    """Modifica le celle specificate direttamente nell'XML del foglio xlsx."""
    tmp = xlsx_path + ".tmp"

    with zipfile.ZipFile(xlsx_path, "r") as zin:
        sheet_path = _find_sheet_path(zin)
        sheet_bytes = zin.read(sheet_path)

        # Registra tutti i namespace per evitare prefissi ns0:
        for _, (prefix, uri) in ET.iterparse(io.BytesIO(sheet_bytes), events=["start-ns"]):
            ET.register_namespace(prefix if prefix else "", uri)

        tree = ET.parse(io.BytesIO(sheet_bytes))
        root = tree.getroot()

        # Determina il namespace principale
        ns = _NS
        tag_match = re.match(r"\{(.+?)\}", root.tag)
        if tag_match:
            ns = tag_match.group(1)

        # Leggi i range uniti (merged cells)
        merge_ranges = []
        mc_elem = root.find(f"{{{ns}}}mergeCells")
        if mc_elem is not None:
            for mc in mc_elem.findall(f"{{{ns}}}mergeCell"):
                ref = mc.get("ref")
                if ref:
                    merge_ranges.append(ref)

        # Risolvi celle unite -> scrivi nella cella principale
        resolved = {}
        for ref, val in cells_to_write.items():
            resolved[_resolve_merge(ref, merge_ranges)] = val

        # Modifica le celle
        sheet_data = root.find(f"{{{ns}}}sheetData")
        for ref, val in resolved.items():
            _, row_num = _parse_cell_ref(ref)

            # Trova la riga
            row_el = None
            for r in sheet_data.findall(f"{{{ns}}}row"):
                if r.get("r") == str(row_num):
                    row_el = r
                    break
            if row_el is None:
                row_el = ET.SubElement(sheet_data, f"{{{ns}}}row")
                row_el.set("r", str(row_num))

            # Trova la cella
            cell_el = None
            for c in row_el.findall(f"{{{ns}}}c"):
                if c.get("r") == ref:
                    cell_el = c
                    break
            if cell_el is None:
                cell_el = ET.SubElement(row_el, f"{{{ns}}}c")
                cell_el.set("r", ref)

            # Rimuovi valore/formula esistenti, mantieni stile
            for child in list(cell_el):
                ltag = child.tag.rsplit("}", 1)[-1]
                if ltag in ("v", "f", "is"):
                    cell_el.remove(child)

            # Rimuovi attributi formula se presenti
            for attr in ("cm", "vm"):
                if attr in cell_el.attrib:
                    del cell_el.attrib[attr]

            # Scrivi come inline string
            cell_el.set("t", "inlineStr")
            is_el = ET.SubElement(cell_el, f"{{{ns}}}is")
            t_el = ET.SubElement(is_el, f"{{{ns}}}t")
            t_el.text = str(val)

        # Serializza XML modificato
        out_buf = io.BytesIO()
        tree.write(out_buf, xml_declaration=True, encoding="UTF-8")
        modified_sheet = out_buf.getvalue()

        # Riscrivi il file xlsx con il foglio modificato
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == sheet_path:
                    zout.writestr(item, modified_sheet)
                else:
                    zout.writestr(item, zin.read(item.filename))

    os.replace(tmp, xlsx_path)


def get_cell_map() -> dict:
    """Restituisce la mappatura celle corrente per visualizzazione nella GUI."""
    return {
        "Cliente (G2)": "cliente",
        "Elemento assiemato (D2)": "posizioni_stringa",
        "N° commessa (D3)": "numero_commessa",
        "Progetto (D4)": "progetto",
        "Data collaudo (G6)": "data_collaudo",
        "Data controllo (G10)": "data_ddt",
        "Data compilazione (I14)": "data_ddt",
    }


if __name__ == "__main__":
    print("=== Mappatura celle Excel ===")
    for label, field in get_cell_map().items():
        print(f"  {label} <- campo PDF: {field}")
