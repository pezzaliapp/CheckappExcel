# CheckappExcel

<p align="center">
  <img src="assets/og-image.png" alt="CheckappExcel" width="600">
</p>

Applicazione Python per confrontare **due o più file Excel/CSV** di listini prodotti — anche con **più fogli** ciascuno — e produrre un **unico file Excel** con **colonne colorate per file**, evidenziando prodotti presenti/mancanti, differenze di prezzo, trasporto e installazione.

La chiave primaria del confronto è sempre il **codice prodotto**.

---

## 🌐 Versione web (nessuna installazione)

👉 **[Usa l'app online](https://www.alessandropezzali.it/CheckappExcel/)**

Carica i file direttamente nel browser e scarica l'Excel confrontato. Tutto avviene lato client: nessun file viene caricato su server.

Se invece preferisci la versione desktop / da integrare in script, leggi le sezioni qui sotto.

---

## ✨ Caratteristiche

- Supporta `.xlsx`, `.xlsm`, `.xls`, `.csv`, `.tsv` (separatori `,` e `;` rilevati automaticamente).
- Accetta **N file** come input (non solo due).
- Gestisce file con **più fogli**: può unirli in un'unica tabella o tenerli distinti.
- **Riconosce automaticamente** i nomi delle colonne (italiano e inglese):
  - codice: `Codice`, `cod.`, `SKU`, `ID`, `Item Code`...
  - descrizione: `Descrizione`, `Desc`, `Nome`, `Description`...
  - prezzo: `Prezzo`, `Prezzo netto`, `Prezzo listino`, `Price`...
  - trasporto: `Trasporto`, `Spese trasporto`, `Spedizione`, `Shipping`...
  - installazione: `Installazione`, `Montaggio`, `Installation`, `Setup`...
- Codici normalizzati: trim, gestione `.0` dei numerici, case-insensitive (configurabile).
- Output Excel con:
  - un blocco di colonne **per ciascun file**, con **colore di intestazione dedicato**;
  - celle **gialle** per i prodotti assenti in quel file;
  - celle **verdi** per i prodotti presenti solo in un file;
  - celle **arancio** sul prezzo quando i valori differiscono fra file;
  - foglio **Riepilogo** con legenda e statistiche;
  - foglio **Mancanti** con solo i codici non presenti ovunque.
- Doppia interfaccia: **CLI** + **GUI Tkinter**.

---

## 📦 Installazione

Richiede Python 3.9+.

```bash
git clone https://github.com/pezzaliapp/CheckappExcel.git
cd CheckappExcel
pip install -r requirements.txt
```

---

## 🚀 Uso rapido

### Da terminale

```bash
python -m checkapp fornitore_a.xlsx fornitore_b.xlsx -o confronto.xlsx
```

Con più file e CSV:

```bash
python -m checkapp listino_A.xlsx listino_B.xlsx listino_C.csv \
    -o risultato.xlsx \
    --labels "Fornitore A" "Fornitore B" "Fornitore C"
```

Opzioni:

| Flag                 | Descrizione                                                          |
|----------------------|----------------------------------------------------------------------|
| `-o, --output`       | Path del file Excel di output (default: `confronto.xlsx`).           |
| `-l, --labels`       | Etichette da mostrare nel report (una per file).                     |
| `--case-sensitive`   | Confronto codici sensibile a maiuscole/minuscole.                    |
| `--no-merge-sheets`  | Non unisce i fogli: ogni foglio diventa una colonna separata.        |

### Interfaccia grafica

```bash
python gui.py
```

![GUI](docs/gui.png)

Dalla GUI puoi:
1. Aggiungere più file.
2. Rinominare l'etichetta che comparirà nel report.
3. Scegliere dove salvare l'output.
4. Avviare il confronto e aprire direttamente la cartella di destinazione.

### Uso come libreria

```python
from checkapp import run_comparison, CompareOptions

result = run_comparison(
    ["listino_A.xlsx", "listino_B.xlsx"],
    output_path="out.xlsx",
    labels=["Fornitore A", "Fornitore B"],
    options=CompareOptions(merge_sheets=True, case_sensitive_codes=False),
)
print(result["stats"])
```

---

## 📊 Struttura dell'Excel prodotto

Il file di output contiene 3 fogli:

### 1. Riepilogo
Conteggio totale dei codici, quanti sono in tutti i file, quanti parzialmente, quanti unici a un file, più la legenda dei colori.

### 2. Confronto
Una riga per ogni **codice distinto** trovato nell'unione dei file, e blocchi di 4 colonne per ciascun file:

```
Codice | Stato | [Fornitore A: Desc | Prezzo | Trasp | Install] | [Fornitore B: ...] | ...
```

- L'intestazione di ciascun blocco è colorata in modo univoco (blu, verde, arancio, viola, ...).
- Le celle **gialle** segnalano che il prodotto non è presente in quel file.
- Le celle **verdi** segnalano che il prodotto è presente **solo in quel file**.
- Le celle **arancio** sulla colonna `Prezzo` segnalano che i prezzi differiscono tra i file.
- I prezzi sono formattati come valuta (`€`).
- Filtri automatici e `freeze panes` già configurati.

### 3. Mancanti
Solo i codici che **non** sono presenti in tutti i file, con una matrice `Sì/No` per file.

---

## 🧪 Esempi inclusi

La cartella `examples/` contiene tre file di prova:

- `fornitore_A.xlsx` — listino semplice (1 foglio).
- `fornitore_B.xlsx` — 2 fogli (`Listino` + `Promozioni`), intestazioni diverse.
- `fornitore_C.csv` — CSV con separatore `;` e colonne in inglese.

Prova:

```bash
python examples/_generate_examples.py   # (ri)genera i file di esempio
python -m checkapp examples/fornitore_A.xlsx examples/fornitore_B.xlsx \
                  examples/fornitore_C.csv -o examples/confronto_esempio.xlsx
```

---

## 🧩 Personalizzazione

Se i tuoi listini usano nomi colonna particolari, puoi estendere la mappa degli alias passando un `CompareOptions` personalizzato:

```python
from checkapp import CompareOptions, run_comparison

opts = CompareOptions(output_path="out.xlsx")
opts.column_aliases["codice"].append("matricola")
opts.column_aliases["prezzo"].append("imponibile")

run_comparison(["a.xlsx", "b.xlsx"], output_path="out.xlsx", options=opts)
```

---

## 🛠️ Sviluppo

```bash
# test
python -m unittest discover -s tests -v
```

Struttura del progetto:

```
CheckappExcel/
├── checkapp/
│   ├── __init__.py
│   ├── __main__.py       # permette `python -m checkapp`
│   ├── comparator.py     # logica di confronto ed export
│   ├── cli.py            # interfaccia a riga di comando
│   └── gui.py            # interfaccia Tkinter
├── examples/             # file di prova
├── tests/                # unit test
├── gui.py                # shortcut per lanciare la GUI
├── requirements.txt
├── LICENSE
└── README.md
```

---

## 📄 Licenza

MIT — vedi [LICENSE](LICENSE).
