# One Piece Card Collection

Gestore locale per collezione **One Piece Card Game**.

Il progetto crea e aggiorna un file Excel con:

- carte dal sito ufficiale Bandai;
- prezzi Cardmarket scaricati automaticamente;
- quantità possedute;
- valore totale;
- dashboard Excel con KPI e grafici;
- confronto prezzo rispetto all'ultimo JSON finale salvato in backup;
- report Top 5 carte in aumento e Top 5 carte in calo;
- dashboard web locale opzionale con Streamlit.

## Struttura progetto

```text
one-piece-card-collection/
  README.md
  requirements.txt
  .gitignore

  src/
    one_piece_common.py
    one_piece_collection_build.py
    one_piece_collection_update_values.py
    one_piece_collection_sync.py
    one_piece_app.py
    README.md

  json/
    .gitkeep

  out/
    .gitkeep

  stg/
    .gitkeep

  logs/
    .gitkeep

  bkp/
    .gitkeep
```

## Installazione

Da terminale, nella root del progetto:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Su macOS/Linux:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Comandi principali

### 1. Creare tutto da zero

```bash
python src/one_piece_collection_build.py
```

Fa scraping del sito Bandai, scarica i JSON Cardmarket, genera Excel e JSON finale.

### 2. Aggiornare solo i valori

```bash
python src/one_piece_collection_update_values.py
```

Mantiene le quantità già inserite e aggiorna prezzi/andamento usando i JSON Cardmarket più recenti.

### 3. Sincronizzare nuove espansioni

```bash
python src/one_piece_collection_sync.py
```

Parte dall'Excel esistente, mantiene le quantità, rilegge Bandai e Cardmarket, aggiorna prezzi e aggiunge nuove carte/espansioni.

### 4. Aprire dashboard web locale

Metodo consigliato:

```bash
streamlit run src/one_piece_app.py
```

Funziona anche così, perché lo script rilancia automaticamente Streamlit:

```bash
python src/one_piece_app.py
```

Apre una dashboard nel browser con:

- KPI valore/quantità;
- grafici per rarità, lingua, trend prezzo ed espansione;
- Top 5 in aumento;
- Top 5 in calo;
- tabella filtrabile;
- pulsanti per lanciare build/update/sync;
- download di Excel e JSON;
- visualizzazione degli ultimi log;
- log live mentre build/update/sync sono in esecuzione.

## Output finali

La cartella `out/` contiene solo i file finali:

```text
out/
  one_piece_collection.xlsx
  one_piece_collection.json
```

L'Excel contiene:

- foglio `Carte`;
- foglio `Dashboard`;
- KPI principali;
- grafici;
- Top 5 aumenti/cali;
- riepiloghi per rarità, colore, lingua, espansione.

## Confronto prezzi

Quando aggiorni o fai sync, il programma confronta il valore corrente con l'ultimo JSON finale salvato in backup:

```text
bkp/<timestamp>/out/one_piece_collection.json
```

Nel foglio `Carte` trovi:

```text
Valore precedente
Variazione valore
Variazione %
Trend prezzo
Data confronto
Fonte confronto
```

I report vengono salvati in `stg/`:

```text
stg/price_trend_report.csv
stg/top_5_aumenti.csv
stg/top_5_cali.csv
```

## Download JSON Cardmarket

Gli script provano a leggere le pagine ufficiali Cardmarket:

- `https://www.cardmarket.com/en/Magic/Data/Product-List`
- `https://www.cardmarket.com/en/Magic/Data/Price-Guide`

Se Cardmarket risponde con `403 Forbidden`, usano automaticamente i fallback S3 pubblici One Piece:

- `products_singles_18.json`
- `products_nonsingles_18.json`
- `price_guide_18.json`

I JSON attivi vengono salvati in:

```text
json/
```

Le fonti usate vengono registrate in:

```text
stg/cardmarket_json_sources.json
```

## Backup

Prima di sovrascrivere file finali o JSON, viene creata una snapshot:

```text
bkp/
  20260701_204512/
    out/
      one_piece_collection.xlsx
      one_piece_collection.json
    json/
      products_singles_18.json
      products_nonsingles_18.json
      price_guide_18.json
    manifest.json
```

I file mantengono il nome originale; il timestamp è nella cartella.

## Log

Ogni esecuzione crea un log in:

```text
logs/
  run_YYYYMMDD_HHMMSS_build.log
  run_YYYYMMDD_HHMMSS_update_values.log
  run_YYYYMMDD_HHMMSS_sync.log
```

I log contengono gli stessi messaggi mostrati a console.

Dalla dashboard Streamlit, quando premi `Crea tutto da zero`, `Aggiorna solo valori` o `Sync nuove espansioni`, il log viene mostrato in tempo reale nella pagina.

Alla fine della run puoi usare `Ricarica dati dashboard` per rileggere Excel/JSON aggiornati.

## GitHub

`.gitignore` ignora i contenuti generati di:

- `json/`
- `out/`
- `stg/`
- `logs/`
- `bkp/`

ma mantiene le cartelle grazie ai `.gitkeep`.

## Note

- Chiudi Excel prima di eseguire gli script.
- Al primo run non esiste ancora un confronto prezzi precedente.
- Dal secondo run in poi la dashboard mostra aumento/calo rispetto al JSON precedente.

## Modifica quantità da Streamlit

Nella dashboard web, la tabella carte è modificabile solo nella colonna `Quantità`.

Procedura:

1. filtra o cerca la carta;
2. modifica il valore nella colonna `Quantità`;
3. premi `Salva quantità`;
4. il programma aggiorna automaticamente:
   - `out/one_piece_collection.json`
   - `out/one_piece_collection.xlsx`
5. prima del salvataggio viene creato un backup automatico dei file finali esistenti.

I campi economici nella dashboard web sono indicati in euro:

```text
Valore (€)
Valore totale (€)
Valore precedente (€)
Variazione valore (€)
```

Nota: le quantità modificate da Streamlit vengono mantenute anche nei successivi `update_values` e `sync`.


## Ultime modifiche

- I JSON Cardmarket non vengono più cercati leggendo le pagine HTML di Cardmarket: gli script usano direttamente gli URL S3 pubblici configurati in `one_piece_common.py`.
- La dashboard usa il termine **Carte** invece di **righe**.
- Ogni build, sync, update valori e salvataggio quantità aggiorna `stg/value_history.csv`.
- Streamlit mostra un grafico **Valore collezione nel tempo** usando lo storico salvato.
- I valori economici sono espressi in euro (€).
