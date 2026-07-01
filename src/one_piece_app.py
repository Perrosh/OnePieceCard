import json
import os
import subprocess
import sys
import time
from pathlib import Path
from collections import deque
from zipfile import BadZipFile

import pandas as pd
import plotly.express as px
import streamlit as st

from one_piece_common import (
    PROJECT_ROOT,
    OUTPUT_XLSX,
    OUTPUT_JSON,
    STG_DIR,
    LOG_DIR,
    JSON_DIR,
    BKP_DIR,
    PRICE_TREND_REPORT_CSV,
    TOP_5_AUMENTI_CSV,
    TOP_5_CALI_CSV,
    FINAL_STG_CSV,
    UPDATED_STG_CSV,
)


def running_inside_streamlit():
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except Exception:
        return False


# Permette di avviare l'app anche con:
#   python src/one_piece_app.py
# In quel caso rilancia automaticamente:
#   python -m streamlit run src/one_piece_app.py
if __name__ == "__main__" and not running_inside_streamlit():
    script_path = Path(__file__).resolve()
    cmd = [sys.executable, "-m", "streamlit", "run", str(script_path)]
    print("Avvio Streamlit:")
    print(" ".join(f'"{x}"' if " " in x else x for x in cmd))
    raise SystemExit(subprocess.call(cmd))

st.set_page_config(
    page_title="One Piece Card Collection",
    page_icon="🏴‍☠️",
    layout="wide",
)

ROOT = Path(PROJECT_ROOT)
SRC = ROOT / "src"


def euro(value):
    try:
        return f"€ {float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "€ 0,00"


def _read_json_collection(json_path: Path):
    with json_path.open("r", encoding="utf-8") as f:
        payload = json.load(f)
    rows = payload.get("cards", payload if isinstance(payload, list) else [])
    generated_at = payload.get("generatedAt", "") if isinstance(payload, dict) else ""
    return pd.DataFrame(rows), str(json_path), generated_at


def _read_excel_collection(xlsx_path: Path):
    try:
        df = pd.read_excel(xlsx_path, sheet_name="Carte")
        return df, str(xlsx_path), ""
    except (BadZipFile, ValueError, OSError) as exc:
        st.warning(
            "Ho trovato il file Excel finale, ma non sembra un .xlsx valido. "
            "Probabilmente una build precedente si è interrotta mentre lo stava creando. "
            f"File ignorato: {xlsx_path} | Errore: {exc}"
        )
        return None


def _read_csv_collection(csv_path: Path):
    try:
        df = pd.read_csv(csv_path, encoding="utf-8-sig")
        return df, str(csv_path), ""
    except Exception as exc:
        st.warning(f"Non riesco a leggere il CSV di fallback {csv_path}: {exc}")
        return None


def load_collection():
    """Carica la collezione per la dashboard web.

    Priorità:
    1. out/one_piece_collection.json, il formato più robusto;
    2. out/one_piece_collection.xlsx, solo se valido;
    3. CSV intermedi in stg/, utili quando la build si è fermata prima di creare l'Excel.
    """
    json_path = Path(OUTPUT_JSON)
    xlsx_path = Path(OUTPUT_XLSX)

    if json_path.exists() and json_path.stat().st_size > 0:
        try:
            return _read_json_collection(json_path)
        except Exception as exc:
            st.warning(f"JSON finale presente ma non leggibile, provo altri file. Errore: {exc}")

    if xlsx_path.exists() and xlsx_path.stat().st_size > 0:
        loaded = _read_excel_collection(xlsx_path)
        if loaded is not None:
            return loaded

    fallback_csvs = [
        Path(FINAL_STG_CSV),
        Path(UPDATED_STG_CSV),
        Path(STG_DIR) / "one_piece_collection_sync_stg.csv",
        Path(STG_DIR) / "one_piece_collection_stg.csv",
    ]
    existing = [p for p in fallback_csvs if p.exists() and p.stat().st_size > 0]
    existing.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    for csv_path in existing:
        loaded = _read_csv_collection(csv_path)
        if loaded is not None:
            st.info(
                "Sto usando un CSV intermedio perché non ho trovato un JSON/Excel finale valido. "
                "Premi 'Crea tutto da zero' dopo aver installato l'ultima versione per rigenerare out/one_piece_collection.xlsx."
            )
            return loaded

    return pd.DataFrame(), "", ""


def normalize_numeric(df, cols):
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            out[c] = 0
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0)
    return out


def run_script(script_name):
    script_path = SRC / script_name
    if not script_path.exists():
        st.error(f"Script non trovato: {script_path}")
        return False

    st.session_state["last_run_script"] = script_name
    st.session_state["last_run_output"] = ""
    st.session_state["last_run_returncode"] = None

    status = st.status(f"Eseguo {script_name}...", expanded=True)
    log_box = st.empty()
    tail_box = st.empty()

    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"

    cmd = [sys.executable, "-u", str(script_path)]
    all_lines = []
    tail_lines = deque(maxlen=250)
    started_at = time.time()

    try:
        process = subprocess.Popen(
            cmd,
            cwd=str(ROOT),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            env=env,
        )

        assert process.stdout is not None
        for line in process.stdout:
            clean = line.rstrip("\n")
            all_lines.append(clean)
            tail_lines.append(clean)

            elapsed = int(time.time() - started_at)
            tail_box.caption(f"Run attiva da {elapsed}s. Ultime {len(tail_lines)} righe mostrate.")
            log_box.code("\n".join(tail_lines), language="text")

        returncode = process.wait()

    except Exception as exc:
        all_lines.append(f"ERRORE AVVIO SCRIPT: {exc}")
        returncode = -1

    output = "\n".join(all_lines)
    st.session_state["last_run_output"] = output
    st.session_state["last_run_returncode"] = returncode

    elapsed = int(time.time() - started_at)
    log_box.code(output[-30000:] if output else "Nessun output prodotto.", language="text")

    if returncode == 0:
        status.update(label=f"{script_name} completato in {elapsed}s.", state="complete", expanded=True)
        st.success(f"{script_name} completato.")
        return True

    status.update(label=f"{script_name} terminato con errore {returncode} dopo {elapsed}s.", state="error", expanded=True)
    st.error(f"{script_name} terminato con errore {returncode}.")
    return False


def latest_log_files(limit=5):
    log_dir = Path(LOG_DIR)
    if not log_dir.exists():
        return []
    files = [p for p in log_dir.glob("*.log") if p.is_file()]
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return files[:limit]


st.title("🏴‍☠️ One Piece Card Collection")
st.caption("Dashboard locale per Excel, JSON, valori Cardmarket, trend prezzo e log di aggiornamento.")

with st.sidebar:
    st.header("Comandi")
    st.write("Lancia gli script senza aprire PyCharm.")

    if st.button("Crea tutto da zero", use_container_width=True):
        st.session_state["pending_script"] = "one_piece_collection_build.py"

    if st.button("Aggiorna solo valori", use_container_width=True):
        st.session_state["pending_script"] = "one_piece_collection_update_values.py"

    if st.button("Sync nuove espansioni", use_container_width=True):
        st.session_state["pending_script"] = "one_piece_collection_sync.py"

    if st.button("Ricarica dati dashboard", use_container_width=True):
        st.rerun()

    st.divider()
    st.write("Cartelle")
    st.code(
        f"root: {ROOT}\njson: {JSON_DIR}\nout: {Path(OUTPUT_XLSX).parent}\nstg: {STG_DIR}\nlogs: {LOG_DIR}\nbkp: {BKP_DIR}",
        language="text",
    )


pending_script = st.session_state.pop("pending_script", None)
if pending_script:
    st.subheader("Esecuzione in corso")
    st.caption("Il log qui sotto viene aggiornato mentre lo script lavora. Se Cardmarket o Bandai sono lenti, almeno vedi il battito del motore.")
    completed_ok = run_script(pending_script)
    if completed_ok:
        st.info("Run completata. La dashboard qui sotto viene ricaricata leggendo i file appena creati.")
    else:
        st.warning("Run terminata con errore. Leggi il log qui sopra: l'ultima traceback è la mappa del tesoro per sistemare il bug.")

if st.session_state.get("last_run_output"):
    with st.expander("Ultimo log esecuzione da Streamlit", expanded=False):
        rc = st.session_state.get("last_run_returncode")
        script = st.session_state.get("last_run_script", "")
        st.caption(f"Script: {script} | Exit code: {rc}")
        st.code(st.session_state["last_run_output"][-30000:], language="text")

collection, source, generated_at = load_collection()

if collection.empty:
    st.warning("Non trovo ancora una collezione. Lancia 'Crea tutto da zero' oppure esegui lo script build da terminale.")
    st.stop()

collection = normalize_numeric(
    collection,
    ["Quantità", "Valore", "Valore totale", "Valore precedente", "Variazione valore", "Variazione %"],
)

# Se il JSON contiene formule/vuoti o manca una colonna, ricalcolo comunque per la dashboard web.
collection["Valore totale"] = collection["Quantità"] * collection["Valore"]

st.info(f"Fonte dati caricata: {source}" + (f" | generato: {generated_at}" if generated_at else ""))

k1, k2, k3, k4, k5, k6 = st.columns(6)
rows = len(collection)
owned_rows = int((collection.get("Quantità", 0) > 0).sum())
total_qty = int(collection.get("Quantità", 0).sum())
total_value = float(collection.get("Valore totale", 0).sum())
priced = int((collection.get("Valore", 0) > 0).sum())
net_delta = float(collection.get("Variazione valore", 0).sum()) if "Variazione valore" in collection.columns else 0

k1.metric("Righe", rows)
k2.metric("Righe possedute", owned_rows)
k3.metric("Quantità totale", total_qty)
k4.metric("Valore totale", euro(total_value))
k5.metric("Con prezzo", priced)
k6.metric("Delta netto", euro(net_delta))

if "Trend prezzo" in collection.columns:
    t1, t2, t3, t4 = st.columns(4)
    t1.metric("In aumento", int((collection["Trend prezzo"] == "In aumento").sum()))
    t2.metric("In calo", int((collection["Trend prezzo"] == "In calo").sum()))
    t3.metric("Stabili", int((collection["Trend prezzo"] == "Stabile").sum()))
    t4.metric("Non confrontate", int(collection["Trend prezzo"].isin(["Nuova / non confrontata", "Nessun confronto"]).sum()))

st.divider()

chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    if "Rarità" in collection.columns:
        rarity = collection.groupby("Rarità", dropna=False).agg(quantita=("Quantità", "sum")).reset_index()
        rarity = rarity[rarity["Rarità"].astype(str).str.strip() != ""]
        if not rarity.empty:
            fig = px.bar(rarity, x="Rarità", y="quantita", title="Quantità possedute per rarità")
            st.plotly_chart(fig, use_container_width=True)

with chart_col2:
    if "Lingua" in collection.columns:
        lang = collection.groupby("Lingua", dropna=False).agg(valore=("Valore totale", "sum")).reset_index()
        lang = lang[lang["Lingua"].astype(str).str.strip() != ""]
        if not lang.empty:
            fig = px.pie(lang, names="Lingua", values="valore", title="Valore posseduto per lingua")
            st.plotly_chart(fig, use_container_width=True)

chart_col3, chart_col4 = st.columns(2)

with chart_col3:
    if "Trend prezzo" in collection.columns:
        trend = collection.groupby("Trend prezzo", dropna=False).size().reset_index(name="righe")
        trend = trend[trend["Trend prezzo"].astype(str).str.strip() != ""]
        if not trend.empty:
            fig = px.bar(trend, x="Trend prezzo", y="righe", title="Andamento prezzi per numero righe")
            st.plotly_chart(fig, use_container_width=True)

with chart_col4:
    if "Espansione" in collection.columns:
        exp = collection.groupby("Espansione", dropna=False).agg(valore=("Valore totale", "sum")).reset_index()
        exp = exp.sort_values("valore", ascending=False).head(15)
        if not exp.empty:
            fig = px.bar(exp, x="valore", y="Espansione", orientation="h", title="Top 15 espansioni per valore posseduto")
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)

st.divider()

up_col, down_col = st.columns(2)

with up_col:
    st.subheader("Top 5 in aumento")
    if "Variazione valore" in collection.columns:
        top_up = collection[collection["Variazione valore"] > 0].sort_values("Variazione valore", ascending=False).head(5)
        cols = [c for c in ["ID Carta", "Nome", "Lingua", "Variante", "Valore precedente", "Valore", "Variazione valore", "Variazione %"] if c in top_up.columns]
        st.dataframe(top_up[cols], use_container_width=True, hide_index=True)
    else:
        st.caption("Nessun confronto disponibile.")

with down_col:
    st.subheader("Top 5 in calo")
    if "Variazione valore" in collection.columns:
        top_down = collection[collection["Variazione valore"] < 0].sort_values("Variazione valore", ascending=True).head(5)
        cols = [c for c in ["ID Carta", "Nome", "Lingua", "Variante", "Valore precedente", "Valore", "Variazione valore", "Variazione %"] if c in top_down.columns]
        st.dataframe(top_down[cols], use_container_width=True, hide_index=True)
    else:
        st.caption("Nessun confronto disponibile.")

st.divider()

st.subheader("Tabella carte")
filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)

filtered = collection.copy()

with filter_col1:
    if "Lingua" in filtered.columns:
        langs = sorted([x for x in filtered["Lingua"].dropna().astype(str).unique() if x.strip()])
        selected_langs = st.multiselect("Lingua", langs, default=[])
        if selected_langs:
            filtered = filtered[filtered["Lingua"].isin(selected_langs)]

with filter_col2:
    if "Espansione" in filtered.columns:
        exps = sorted([x for x in filtered["Espansione"].dropna().astype(str).unique() if x.strip()])
        selected_exps = st.multiselect("Espansione", exps, default=[])
        if selected_exps:
            filtered = filtered[filtered["Espansione"].isin(selected_exps)]

with filter_col3:
    if "Rarità" in filtered.columns:
        rarities = sorted([x for x in filtered["Rarità"].dropna().astype(str).unique() if x.strip()])
        selected_rarities = st.multiselect("Rarità", rarities, default=[])
        if selected_rarities:
            filtered = filtered[filtered["Rarità"].isin(selected_rarities)]

with filter_col4:
    owned_only = st.checkbox("Solo possedute")
    if owned_only and "Quantità" in filtered.columns:
        filtered = filtered[filtered["Quantità"] > 0]

search = st.text_input("Cerca per nome o ID")
if search:
    s = search.lower().strip()
    mask = pd.Series(False, index=filtered.index)
    for col in ["ID Carta", "Nome", "Variante"]:
        if col in filtered.columns:
            mask = mask | filtered[col].astype(str).str.lower().str.contains(s, na=False)
    filtered = filtered[mask]

show_cols = [c for c in [
    "ID Carta", "Espansione", "Numero", "Nome", "Lingua", "Variante", "Rarità", "Color", "Quantità", "Valore", "Valore totale",
    "Valore precedente", "Variazione valore", "Variazione %", "Trend prezzo"
] if c in filtered.columns]

st.dataframe(filtered[show_cols], use_container_width=True, hide_index=True)

st.divider()

download_col1, download_col2, log_col = st.columns(3)

with download_col1:
    xlsx = Path(OUTPUT_XLSX)
    if xlsx.exists():
        st.download_button(
            "Scarica Excel",
            data=xlsx.read_bytes(),
            file_name=xlsx.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

with download_col2:
    js = Path(OUTPUT_JSON)
    if js.exists():
        st.download_button(
            "Scarica JSON finale",
            data=js.read_bytes(),
            file_name=js.name,
            mime="application/json",
            use_container_width=True,
        )

with log_col:
    logs = latest_log_files()
    if logs:
        selected = st.selectbox("Ultimi log", logs, format_func=lambda p: p.name)
        with st.expander("Mostra log selezionato"):
            st.code(selected.read_text(encoding="utf-8", errors="replace")[-20000:], language="text")
    else:
        st.caption("Nessun log trovato.")
