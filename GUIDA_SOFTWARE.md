# Guida Completa - Gestione Qualita: Compilazione Automatica Excel da PDF DOP

## CM SRL - Controllo Prefabbricazione

---

## 1. Cos'e questo software

Questo software e uno strumento desktop per la **gestione della qualita nel settore della prefabbricazione**. Automatizza la compilazione di fogli Excel di controllo prefabbricazione partendo da certificati PDF DOP (Dichiarazione di Prestazione CE).

### Funzionalita principali

- **Estrazione automatica dati da PDF DOP**: legge i certificati di prestazione e ne estrae numero DOP, numero DDT, data DDT, tipo prodotto, codici posizione, fabbricante e norma di riferimento.
- **Compilazione automatica Excel**: popola il template Excel "Controllo Prefabbricazione" con i dati estratti, creando un foglio numerato (001, 002, 003...) per ogni PDF.
- **Elaborazione batch**: carica e processa piu PDF contemporaneamente.
- **Ordinamento cronologico**: i fogli vengono numerati in ordine di data DDT (dal piu vecchio al piu recente).
- **Filtraggio con Distinta Spedizione**: incrocia i codici posizione con la distinta per includere solo quelli pertinenti.
- **Integrazione con Marcature**: estrae le date di collaudo piu recenti dal file Excel delle marcature.
- **Calcolo automatico date**: calcola le date di ispezione e controllo (saltando i weekend).

---

## 2. Requisiti di sistema

- **Python 3.x**
- **Librerie richieste** (installabili con `pip install -r requirements.txt`):
  - `openpyxl >= 3.1.0` - gestione file Excel .xlsx
  - `PyPDF2 >= 3.0.0` - estrazione testo da PDF
  - `xlrd >= 2.0.1` - lettura file Excel legacy .xls

---

## 3. Installazione

```bash
# 1. Clonare il repository
git clone <url-repository>
cd GestioneQualit-

# 2. Installare le dipendenze
pip install -r requirements.txt

# 3. Avviare l'applicazione
python app.py
```

---

## 4. Interfaccia utente

L'applicazione si presenta con una finestra organizzata in queste sezioni:

### 4.1 Intestazione
Titolo dell'applicazione: "Compilazione Automatica Excel da PDF DOP" con sottotitolo "CM SRL - Controllo Prefabbricazione".

### 4.2 Selezione File (sezione superiore)

| Campo | Descrizione |
|-------|-------------|
| **PDF DOP** | Lista dei file PDF caricati. Si aggiungono tramite "Aggiungi PDF..." e si svuotano con "Cancella lista". |
| **Template Excel** | Il file template `.xlsx` da usare come base (default: `CONTROLLO_PREFABBRICAZIONE_.xlsx`). |

### 4.3 Dati Commessa (sezione centrale)

| Campo | Descrizione |
|-------|-------------|
| **N. Commessa** | Numero identificativo della commessa |
| **Progetto** | Nome del progetto |
| **Cliente** | Nome del cliente |
| **Responsabile** | Nome del responsabile/supervisore di progetto |
| **Resp. Saldatura** | Nome del responsabile della saldatura |
| **Excel Marcature** | File Excel contenente le marcature con date di collaudo (opzionale) |
| **Distinta Spedizione** | File con la distinta di spedizione per filtrare i codici posizione (opzionale) |

### 4.4 Pulsanti di azione

- **"1. Estrai dati dai PDF"** - Avvia l'estrazione dei dati dai PDF caricati
- **"2. Compila Excel e Salva"** - Genera i fogli Excel compilati

### 4.5 Anteprima
Area di testo che mostra i dati estratti da ogni PDF prima della compilazione.

### 4.6 Barra di stato
In basso, mostra messaggi di avanzamento e conferma.

---

## 5. Procedura operativa passo-passo

### Passo 1: Preparare i file

Assicurarsi di avere:
- I **file PDF DOP** (Dichiarazioni di Prestazione CE) da processare
- Il **template Excel** (`CONTROLLO_PREFABBRICAZIONE_.xlsx`) presente nella cartella del programma
- (Opzionale) Il file **Excel Marcature** con le date di collaudo
- (Opzionale) La **Distinta Spedizione** per filtrare i codici posizione

### Passo 2: Avviare il software

```bash
python app.py
```

### Passo 3: Caricare i PDF

1. Cliccare **"Aggiungi PDF..."**
2. Selezionare uno o piu file PDF DOP dalla finestra di dialogo
3. I file appariranno nella lista nella parte superiore dell'interfaccia
4. Per ricominciare, usare **"Cancella lista"**

### Passo 4: Selezionare il template Excel

- Se il template `CONTROLLO_PREFABBRICAZIONE_.xlsx` e nella stessa cartella, viene rilevato automaticamente
- Altrimenti, cliccare **"Sfoglia..."** accanto al campo del template Excel per selezionarlo

### Passo 5: Compilare i dati della commessa

Inserire nei campi di testo:
- **N. Commessa**: es. `COM-2025-001`
- **Progetto**: es. `Struttura Prefabbricata Nord`
- **Cliente**: es. `Azienda XYZ S.r.l.`
- **Responsabile**: es. `Mario Rossi`
- **Resp. Saldatura**: es. `Luigi Bianchi`

### Passo 6 (Opzionale): Caricare file aggiuntivi

- **Excel Marcature**: cliccare "Sfoglia..." per selezionare il file con le marcature e le date di collaudo
  - Il software legge la Colonna A (codice posizione) e la Colonna D (data) per trovare la data di collaudo piu recente
- **Distinta Spedizione**: cliccare "Sfoglia..." per selezionare la distinta
  - Usata per filtrare i codici posizione: solo quelli presenti nella distinta verranno inclusi nel foglio Excel

### Passo 7: Estrarre i dati dai PDF

1. Cliccare **"1. Estrai dati dai PDF"**
2. Il software leggera ogni PDF e ne estrarra:
   - Numero DOP (es. `0474-CPR-2548`)
   - Numero e data DDT
   - Tipo di prodotto (es. `ASCENSORE B`)
   - Codici posizione (es. `T41, T42, A56`)
   - Fabbricante
   - Norma di riferimento (es. `EN 1090-2:2018+A1:2020`)
3. I risultati vengono mostrati nell'area **Anteprima**
4. I PDF vengono ordinati per data DDT (dal piu vecchio al piu recente)

**Controllare l'anteprima** per verificare che i dati siano stati estratti correttamente prima di procedere.

### Passo 8: Compilare l'Excel e salvare

1. Cliccare **"2. Compila Excel e Salva"**
2. Selezionare la cartella di destinazione dalla finestra di dialogo
3. Il software generera un file Excel per ogni PDF, con:
   - Numerazione progressiva dei fogli (Scheda N. 001, 002, 003...)
   - Tutti i dati estratti inseriti nelle celle appropriate
   - Date di collaudo e controllo calcolate automaticamente
   - Formattazione, immagini e formule del template preservate

---

## 6. Mappatura celle Excel

Il template viene compilato con i seguenti riferimenti:

| Cella | Contenuto |
|-------|-----------|
| **C1** | Numero scheda (001, 002, 003...) |
| **D2** | Codici posizione (es. `T41-T42-T43`) |
| **D3** | Numero commessa |
| **D4** | Nome progetto |
| **E6-E9, E11-E12** | Nome responsabile |
| **E10** | Responsabile saldatura |
| **G2** | Nome cliente |
| **G6** | Data collaudo (da Excel Marcature, la piu recente) |
| **G7** | Uguale a G6 |
| **G8** | G6 + 2 giorni lavorativi (salta weekend) |
| **G9** | G10 - 2 giorni lavorativi (salta weekend) |
| **G10** | Data controllo (= data DDT) |
| **G11** | Copia di G10 |
| **G12** | Copia di G10 |
| **I11** | Riferimento DDT (formato: `DDT <numero>`) |
| **I14** | Data compilazione (= data DDT) |

---

## 7. Logica di calcolo delle date

Il software calcola automaticamente le date di ispezione:

```
G6  = Data collaudo (estratta dal file Marcature, la piu recente tra quelle corrispondenti)
G7  = G6 (stessa data)
G8  = G6 + 2 giorni (se cade di sabato +2, se domenica +1)
G9  = G10 - 2 giorni (se cade di sabato -1, se domenica -2)
G10 = Data DDT (data del documento di trasporto)
```

Le date devono rispettare l'ordine: **G6 <= G7 < G8 < G9 < G10**

---

## 8. Logica di filtraggio

### Filtraggio con Distinta Spedizione
- Il software legge tutti i codici dalla Colonna A di tutti i fogli della distinta
- I codici vengono normalizzati (maiuscolo, senza spazi/punteggiatura)
- Solo i codici posizione presenti nella distinta vengono inclusi nel foglio Excel
- Se dopo il filtraggio non restano codici, il foglio per quel PDF viene saltato (la cella D2 sarebbe vuota)

### Integrazione Marcature
- Legge Colonna A (codice) e Colonna D (data nel formato `1 del 07/04/25`)
- Trova la data piu recente tra tutte le marcature che corrispondono ai codici posizione del PDF
- Questa data viene usata come "Data collaudo" (G6)

---

## 9. Architettura del software

Il progetto e composto da 3 moduli:

```
app.py           -> Interfaccia grafica (Tkinter) e orchestrazione
pdf_parser.py    -> Estrazione dati dai PDF tramite regex
excel_filler.py  -> Compilazione template Excel (manipolazione XML diretta)
```

### Perche manipolazione XML diretta?
Il template Excel viene modificato a livello XML (il formato .xlsx e un archivio ZIP contenente file XML). Questo approccio preserva:
- Formattazione originale
- Immagini inserite
- Formule
- Celle unite
- Namespace XML

---

## 10. Risoluzione problemi

| Problema | Soluzione |
|----------|----------|
| Dati non estratti dal PDF | Verificare che il PDF sia un DOP valido con testo selezionabile (non scansione immagine) |
| Codici posizione mancanti | Controllare che i codici siano nel formato `LETTERA + NUMERI` (es. T41, A56) |
| Foglio Excel vuoto | Verificare che i codici posizione siano presenti nella Distinta Spedizione |
| Date non calcolate | Assicurarsi di aver caricato il file Excel Marcature con i dati corretti |
| Template non trovato | Posizionare `CONTROLLO_PREFABBRICAZIONE_.xlsx` nella stessa cartella di `app.py` |
| Errore di importazione | Eseguire `pip install -r requirements.txt` per installare le dipendenze |

---

## 11. Esempio di flusso completo

```
1. Avvio: python app.py
2. Carico 5 PDF DOP
3. Inserisco: Commessa=COM-001, Progetto=Ponte Nord, Cliente=ABC Srl
4. Inserisco: Responsabile=M. Rossi, Resp. Saldatura=L. Bianchi
5. Carico Excel Marcature e Distinta Spedizione
6. Clicco "1. Estrai dati dai PDF" -> verifico anteprima
7. Clicco "2. Compila Excel e Salva" -> scelgo cartella output
8. Risultato: 5 file Excel (o meno se filtrati), numerati 001-005
   con tutti i dati compilati e le date calcolate
```
