# One Piece Card Collection

Gestore locale per collezione **One Piece Card Game**.

Il progetto crea e aggiorna un file Excel con:

- carte dal sito ufficiale Bandai EN e dal catalogo ufficiale JP/Asia per le rarità JP;
- prezzi Cardmarket scaricati automaticamente;
- quantità possedute;
- valore totale;
- dashboard Excel con KPI e grafici;
- confronto prezzo rispetto all'ultimo JSON finale salvato in backup;
- report e grafici configurabili per le carte in aumento/calo;
- dashboard web locale con Streamlit, temi selezionabili, ricerca semplice/avanzata e gestione carte modificabile;
- storico del valore collezione nel tempo in `stg/value_history.csv`.

## Struttura progetto

```text
one-piece-card-collection/
  README.md
  requirements.txt
  .gitignore

  .streamlit/
    config.toml

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
- Top X in aumento/calo configurabile dall’utente;
- dashboard organizzata in schede/indice: Panoramica, Top valore, Trend prezzi, Gestione carte, Domande database, File e log;
- ricerca principale con i campi più usati e pannello di ricerca avanzata;
- scelta del tema dashboard in alto a destra, delle colonne visibili e del loro ordine nella gestione carte;
- tabella filtrabile e modificabile in tutti i campi visibili, inclusi campi Cardmarket come `Cardmarket idProduct`;
- aggiunta manuale di nuove carte tramite pulsante che crea una riga compilabile;
- eliminazione carte tramite selezione righe in una tabella compatta dedicata;
- avviso visibile quando ci sono modifiche non salvate;
- salvataggio con stato grande e visibile nella pagina;
- rimozione delle righe doppie identiche;
- pulsanti per lanciare build/update/sync;
- download di Excel e JSON;
- visualizzazione degli ultimi log;
- log live mentre build/update/sync sono in esecuzione.


## Organizzazione dashboard Streamlit

### Tema dashboard

Il tema si sceglie dal selettore compatto in alto a destra della dashboard. Streamlit non permette alle app di aggiungere voci personalizzate al menu nativo dei tre puntini del browser/app, quindi il selettore è integrato nella pagina.

Temi disponibili:

- Predefinito
- Chiaro
- Scuro
- Mare
- Wanted
- One Piece


La dashboard web è organizzata in schede che funzionano da indice rapido:

- **Panoramica**: KPI principali e grafici generali;
- **Top valore**: Top X carte per valore, con numero scelto dall’utente;
- **Trend prezzi**: Top X aumenti/cali, con numero scelto dall’utente, e storico valore;
- **Gestione carte**: ricerca semplice/avanzata, scelta colonne, modifica tabella, aggiunta righe, eliminazione carte selezionate da tabella dedicata e rimozione doppioni identici;
- **Domande database**: domande locali sulla collezione, senza chiamate esterne;
- **File e log**: download finali e consultazione log.

La sezione **Trend prezzi** indica una sola volta che i confronti sono calcolati sulle carte possedute (`Quantità > 0`), poi i grafici usano titoli più puliti come `Top X in aumento` e `Top X in calo`.


## Gestione carte da Streamlit

Nella scheda **Gestione carte** puoi lavorare senza aprire direttamente Excel:

- in alto a destra nella dashboard puoi scegliere il tema: Predefinito, Chiaro, Scuro, Mare, Wanted o One Piece;

- in alto trovi la ricerca principale: testo, lingua, espansione e solo possedute;
- il pannello **Ricerca avanzata** contiene rarità, tipo carta, colore, variante, trend prezzo, valore minimo, quantità minima e limite carte visualizzate;
- il pannello **Colonne visibili e ordine** permette di scegliere quali colonne vedere e in quale ordine; tutti i campi visibili nella tabella sono modificabili;
- il pulsante **Aggiungi riga** inserisce una riga vuota nella tabella, da compilare nei campi visibili;
- per inserire tutti i campi, usa il preset completo nelle colonne visibili e poi aggiungi la riga;
- per eliminare carte, apri **Elimina carte**, seleziona le righe nella tabella compatta e premi **Elimina carte selezionate**;
- se modifichi un campo, aggiungi una riga o selezioni carte da eliminare, appare un avviso di modifiche non salvate;
- il salvataggio mostra uno stato visibile nella pagina, crea backup automatico e aggiorna sia JSON sia Excel;
- il pulsante **Elimina doppioni identici** rimuove le righe uguali in tutto e per tutto, mantenendo la prima occorrenza.

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
- Top aumenti/cali;
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

Gli script non cercano più i link nelle pagine HTML di Cardmarket.
Usano direttamente i JSON S3 pubblici One Piece:

- `https://downloads.s3.cardmarket.com/productCatalog/productList/products_singles_18.json`
- `https://downloads.s3.cardmarket.com/productCatalog/productList/products_nonsingles_18.json`
- `https://downloads.s3.cardmarket.com/productCatalog/priceGuide/price_guide_18.json`

I JSON attivi vengono salvati in:

```text
json/
```

Le fonti usate vengono registrate in:

```text
stg/cardmarket_json_sources.json
```

## Catalogo JP e rarità diverse da EN

Il catalogo EN viene letto da:

```text
https://en.onepiece-cardgame.com/cardlist/
```

Il catalogo JP/Asia viene letto da:

```text
https://www.onepiece-cardgame.com/cardlist/
```

Le rarità JP vengono ricavate dal raw ufficiale JP, senza override hardcoded.
Se lo stesso ID carta compare più volte nel catalogo JP, per esempio una stampa `R` e una stampa `TR`, lo script sceglie la rarità con priorità più alta trovata nel sito ufficiale e compila anche le colonne diagnostiche:

```text
Rarità JP ufficiale
Rarità JP candidate
Fonte rarità JP
```

## Modifica dati da Streamlit

Nella tabella carte di Streamlit puoi modificare tutti i campi visibili. I campi Cardmarket, inclusi `Cardmarket idProduct`, sono trattati come campi editabili quando sono mostrati nella tabella.
Dopo aver premuto `Salva modifiche`, il programma aggiorna:

```text
out/one_piece_collection.json
out/one_piece_collection.xlsx
```

Prima del salvataggio viene creato un backup automatico.
Il campo `Valore totale` viene ricalcolato da `Quantità × Valore`.

Nella scheda `Gestione carte` puoi anche:

- inserire una nuova carta manuale compilando ID, nome, lingua, rarità, quantità e valore;
- eliminare carte selezionandole da una tabella compatta dedicata;
- salvare subito JSON ed Excel finali con backup automatico.

## Storico valore

Ogni build, update, sync o salvataggio da Streamlit aggiunge un punto a:

```text
stg/value_history.csv
```

La dashboard Streamlit mostra il grafico `Valore collezione nel tempo (€)`.

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

## Note JP e storico valore

Il catalogo inglese e quello giapponese possono avere rarità diverse. Il progetto scarica/cachea il catalogo JP in `stg/bandai_cards_jp_raw.csv` e applica le rarità JP solo alle righe `Lingua = JP`, senza override hardcoded.

Ogni build/update/sync e ogni salvataggio da Streamlit aggiunge un punto in:

```text
stg/value_history.csv
```

Streamlit mostra il grafico `Valore collezione nel tempo (€)`.

I JSON Cardmarket non vengono più cercati nelle pagine HTML: vengono scaricati direttamente dagli URL S3 configurati in `one_piece_common.py`.


## Novità dashboard Streamlit

### Top X carte per valore

Nella dashboard puoi scegliere quante carte mostrare nella classifica, per esempio Top 5, Top 10 o Top 25.
Puoi ordinare per:

- **Valore posseduto (€)**: usa `Quantità × Valore`;
- **Valore singola carta (€)**: usa il valore unitario della carta.

Di default la classifica considera solo le carte possedute. Puoi includere anche le carte non possedute con l'apposita checkbox.

### Domande sul database

La dashboard contiene una sezione per fare domande testuali sulla collezione. Non usa servizi esterni e non invia dati fuori dal PC: usa un interprete locale basato su regole e Pandas.

Esempi supportati localmente:

```text
qual è la carta più costosa dell'espansione OP03 tra le carte che possiedo?
top 10 carte OP16 per valore posseduto
quali carte JP possiedo che valgono di più?
quante carte SEC possiedo?
top 5 carte in aumento
valore totale OP03
```

La dashboard capisce filtri come espansione (`OP03`, `OP-03`, `ST10`), lingua (`JP`, `EN`), rarità (`SEC`, `SR`, `TR`), carte possedute (`Quantità > 0`), valore unitario, valore posseduto, aumenti e cali.
