# Script

Esegui dalla root del progetto:

```bash
python src/one_piece_collection_build.py
python src/one_piece_collection_update_values.py
python src/one_piece_collection_sync.py
streamlit run src/one_piece_app.py
```

## File

- `one_piece_common.py`: funzioni condivise, dashboard Excel, backup, log, download JSON Cardmarket.
- `one_piece_collection_build.py`: crea tutto da zero.
- `one_piece_collection_update_values.py`: aggiorna solo valori/prezzi, mantenendo quantità.
- `one_piece_collection_sync.py`: rilegge Bandai + Cardmarket, mantiene quantità e aggiunge nuove carte/espansioni.
- `one_piece_app.py`: dashboard web locale Streamlit.
