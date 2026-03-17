"""
Modulo per l'estrazione dati da PDF DOP (Dichiarazione di Prestazione).

Estrae i dati chiave dal documento DOP in formato PDF:
- Numero DOP (es. 0474-CPR-2548)
- Numero DDT e data
- Codici prodotto / posizioni
- Tipo prodotto (es. ASCENSORE B)
"""

import re
from PyPDF2 import PdfReader


def extract_dop_data(pdf_path: str) -> dict:
    """
    Estrae i dati principali dal PDF DOP.

    Args:
        pdf_path: Percorso del file PDF DOP.

    Returns:
        Dizionario con i dati estratti:
        - numero_dop: Numero certificato DOP
        - numero_ddt: Numero DDT
        - data_ddt: Data del DDT
        - nr_riferimento: Numero riferimento (es. 003/25)
        - tipo_prodotto: Tipo prodotto (es. ASCENSORE B)
        - posizioni: Lista codici posizione (es. ['T41', 'T42', ...])
        - posizioni_stringa: Stringa formattata delle posizioni
        - fabbricante: Nome fabbricante
        - norma: Norma armonizzata
    """
    data = {
        "numero_dop": "",
        "numero_ddt": "",
        "data_ddt": "",
        "nr_riferimento": "",
        "tipo_prodotto": "",
        "posizioni": [],
        "posizioni_stringa": "",
        "fabbricante": "",
        "norma": "",
    }

    reader = PdfReader(pdf_path)
    full_text = ""
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            full_text += page_text + "\n"

    if not full_text.strip():
        raise ValueError("Il PDF non contiene testo estraibile.")

    # Estrai numero DOP (es. 0474-CPR-2548)
    dop_match = re.search(r"(\d{4}-CPR-\d+)", full_text)
    if dop_match:
        data["numero_dop"] = dop_match.group(1)

    # Estrai riferimento DDT e data
    # Pattern: Nr. 003/25 rif. DDT N° 140 DEL 09/05/2025
    ddt_match = re.search(
        r"Nr\.\s*(\d+/\d+)\s*rif\.\s*DDT\s*N[°o]?\s*(\d+)\s*DEL\s*(\d{2}/\d{2}/\d{2,4})",
        full_text,
        re.IGNORECASE,
    )
    if ddt_match:
        data["nr_riferimento"] = ddt_match.group(1)
        data["numero_ddt"] = ddt_match.group(2)
        data["data_ddt"] = ddt_match.group(3)

    # Estrai tipo prodotto (es. ASCENSORE B)
    # Cerca dopo "prodotto-tipo" nella riga 1 della tabella
    tipo_match = re.search(
        r"prodotto[\s\-]*tipo\s+(.+?)(?:\n|HEA|IPE|ANGOLARI)",
        full_text,
        re.IGNORECASE,
    )
    if tipo_match:
        data["tipo_prodotto"] = tipo_match.group(1).strip()

    # Estrai posizioni (T41, T42, T43, A56, T28, T29, T30, T27, T31, T32, T34, T35)
    # Cerca tutti i codici posizione nel formato lettera+numeri (T41, A56, etc.)
    posizioni = []
    # Cerca specificamente nella sezione del codice di identificazione
    codice_section = re.search(
        r"(?:prodotto[\s\-]*tipo|identificazione)(.*?)(?:Usi previsti|usi previsti)",
        full_text,
        re.DOTALL | re.IGNORECASE,
    )
    if codice_section:
        section_text = codice_section.group(1)
        # Trova tutti i pattern POS. Txx o posizioni singole
        pos_matches = re.findall(r"POS\.?\s*([A-Z]\d{1,3})", section_text)
        posizioni = list(dict.fromkeys(pos_matches))  # rimuovi duplicati mantenendo ordine

    if not posizioni:
        # Fallback: cerca in tutto il testo
        pos_matches = re.findall(r"POS\.?\s*([A-Z]\d{1,3})", full_text)
        posizioni = list(dict.fromkeys(pos_matches))

    data["posizioni"] = posizioni
    # Formatta come nell'Excel: T41-T42-T43-A56-T28-T29-T30 T27-T31-T32-T34-T35
    data["posizioni_stringa"] = _format_posizioni(posizioni)

    # Estrai fabbricante
    fabb_match = re.search(
        r"Fabbricante\s+(.+?)(?:\n.*?){0,2}(?:\d{5}|\nMandatario)",
        full_text,
        re.DOTALL | re.IGNORECASE,
    )
    if fabb_match:
        data["fabbricante"] = fabb_match.group(1).strip()

    # Estrai norma armonizzata
    norma_match = re.search(r"(EN\s*1090[\s\-:]+\d{4}(?:\+A\d:\d{4})?)", full_text)
    if norma_match:
        data["norma"] = norma_match.group(1).strip()

    return data


def _format_posizioni(posizioni: list) -> str:
    """
    Formatta la lista di posizioni in una stringa leggibile.
    Raggruppa per tipo di profilo (HEA 600, HEA 240, ANGOLARI, IPE, HEA 140).
    """
    if not posizioni:
        return ""
    return "-".join(posizioni)


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        result = extract_dop_data(sys.argv[1])
        print("=== Dati estratti dal PDF DOP ===")
        for key, value in result.items():
            print(f"  {key}: {value}")
    else:
        print("Uso: python pdf_parser.py <percorso_pdf>")
