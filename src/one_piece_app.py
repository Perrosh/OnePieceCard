import json
import os
import subprocess
import sys
import time
from pathlib import Path
from collections import deque

import pandas as pd
import plotly.express as px
import streamlit as st
import requests

from one_piece_common import (
    PROJECT_ROOT,
    OUTPUT_XLSX,
    OUTPUT_JSON,
    STG_DIR,
    LOG_DIR,
    JSON_DIR,
    BKP_DIR,
    PRICE_SOURCE_COLUMN,
    backup_file,
    create_collection_workbook_with_dashboard,
    PRICE_TREND_REPORT_CSV,
    TOP_5_AUMENTI_CSV,
    TOP_5_CALI_CSV,
    VALUE_HISTORY_CSV,
    append_value_history,
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
CONFIG_DIR = ROOT / "config"
OPENAI_TOKEN_FILE = CONFIG_DIR / "openai_api_key.txt"
OPENAI_MODEL_FILE = CONFIG_DIR / "openai_model.txt"
DEFAULT_OPENAI_MODEL = "gpt-4o-mini"


def euro(value):
    try:
        return f"€ {float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "€ 0,00"


def load_collection():
    json_path = Path(OUTPUT_JSON)
    xlsx_path = Path(OUTPUT_XLSX)

    if json_path.exists():
        with json_path.open("r", encoding="utf-8") as f:
            payload = json.load(f)
        rows = payload.get("cards", payload if isinstance(payload, list) else [])
        df = pd.DataFrame(rows)
        source = str(json_path)
        generated_at = payload.get("generatedAt", "") if isinstance(payload, dict) else ""
        return df, source, generated_at

    if xlsx_path.exists():
        df = pd.read_excel(xlsx_path, sheet_name="Carte")
        return df, str(xlsx_path), ""

    return pd.DataFrame(), "", ""


def normalize_numeric(df, cols):
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            out[c] = 0
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0)
    return out


def save_collection_from_streamlit(df, source="streamlit_manual_edit"):
    """Salva modifiche manuali da Streamlit su JSON ed Excel finale.

    La modifica manuale deve essere sicura: prima copio i finali esistenti in bkp/,
    poi scrivo JSON ed Excel tramite file temporanei e sostituzione atomica.
    """
    data = df.copy()
    for col in ["Quantità", "Valore"]:
        if col not in data.columns:
            data[col] = 0
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

    data["Quantità"] = data["Quantità"].clip(lower=0).round(0).astype(int)
    data["Valore"] = data["Valore"].astype(float)
    data["Valore totale"] = data["Quantità"] * data["Valore"]

    # Backup dei finali attuali prima della modifica manuale.
    backup_file(OUTPUT_JSON, move=False)
    backup_file(OUTPUT_XLSX, move=False)

    json_path = Path(OUTPUT_JSON)
    json_tmp = json_path.with_name(json_path.stem + ".tmp" + json_path.suffix)
    payload = {
        "generatedAt": pd.Timestamp.now().isoformat(timespec="seconds"),
        "source": source,
        "priceSourceColumn": PRICE_SOURCE_COLUMN,
        "rows": len(data),
        "cards": data.where(pd.notna(data), "").to_dict(orient="records"),
    }
    with json_tmp.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(json_tmp, json_path)

    xlsx_path = Path(OUTPUT_XLSX)
    xlsx_tmp = xlsx_path.with_name(xlsx_path.stem + ".tmp" + xlsx_path.suffix)
    if xlsx_tmp.exists():
        xlsx_tmp.unlink()
    create_collection_workbook_with_dashboard(data, str(xlsx_tmp))
    os.replace(xlsx_tmp, xlsx_path)
    append_value_history(data, source)


def money_column(label):
    return st.column_config.NumberColumn(label, format="€ %.2f")


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


def ensure_config_dir():
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    gitkeep = CONFIG_DIR / ".gitkeep"
    if not gitkeep.exists():
        gitkeep.write_text("", encoding="utf-8")


def load_saved_token():
    try:
        if OPENAI_TOKEN_FILE.exists():
            return OPENAI_TOKEN_FILE.read_text(encoding="utf-8").strip()
    except Exception:
        return ""
    return ""


def save_token(token):
    ensure_config_dir()
    token = str(token or "").strip()
    if token:
        OPENAI_TOKEN_FILE.write_text(token, encoding="utf-8")
    elif OPENAI_TOKEN_FILE.exists():
        OPENAI_TOKEN_FILE.unlink()


def load_saved_model():
    try:
        if OPENAI_MODEL_FILE.exists():
            return OPENAI_MODEL_FILE.read_text(encoding="utf-8").strip() or DEFAULT_OPENAI_MODEL
    except Exception:
        return DEFAULT_OPENAI_MODEL
    return DEFAULT_OPENAI_MODEL


def save_model(model):
    ensure_config_dir()
    model = str(model or DEFAULT_OPENAI_MODEL).strip() or DEFAULT_OPENAI_MODEL
    OPENAI_MODEL_FILE.write_text(model, encoding="utf-8")


def build_ai_context(df, question, max_rows=250):
    data = df.copy()
    for col in ["Quantità", "Valore", "Valore totale", "Variazione valore"]:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

    q = str(question or "").lower()
    candidates = pd.DataFrame(columns=data.columns)

    # Se nella domanda compare una espansione tipo OP03 / OP-03 / ST10, filtro prima quella.
    import re
    patterns = set()
    for match in re.findall(r"\b(OP|ST|EB|PRB)[- ]?(\d{1,2})\b", q, flags=re.IGNORECASE):
        prefix, number = match
        compact = f"{prefix.upper()}{int(number):02d}"
        dashed = f"{prefix.upper()}-{int(number):02d}"
        patterns.update({compact.lower(), dashed.lower()})

    if patterns:
        mask = pd.Series(False, index=data.index)
        for col in ["ID Carta", "Espansione", "Set", "Codice", "Nome"]:
            if col in data.columns:
                text = data[col].astype(str).str.lower().str.replace("-", "", regex=False)
                for pat in patterns:
                    mask = mask | text.str.contains(pat.replace("-", ""), na=False)
        candidates = data[mask].copy()

    # Altrimenti provo a cercare parole della domanda nei campi testuali più utili.
    if candidates.empty:
        words = [w for w in re.findall(r"[a-zA-Z0-9.\-]{3,}", q) if w not in {"qual", "quale", "carta", "carte", "costosa", "costose", "prezzo", "valore", "della", "del", "set"}]
        if words:
            mask = pd.Series(False, index=data.index)
            for col in ["ID Carta", "Nome", "Espansione", "Rarità", "Lingua", "Variante", "Color", "Tipo carta"]:
                if col in data.columns:
                    text = data[col].astype(str).str.lower()
                    for w in words:
                        mask = mask | text.str.contains(re.escape(w.lower()), na=False)
            candidates = data[mask].copy()

    if candidates.empty:
        candidates = data.copy()

    # Tengo le righe più informative, dando priorità a possedute e valore alto.
    sort_cols = [c for c in ["Quantità", "Valore totale", "Valore"] if c in candidates.columns]
    if sort_cols:
        candidates = candidates.sort_values(sort_cols, ascending=[False] * len(sort_cols))

    cols = [c for c in [
        "ID Carta", "Espansione", "Numero", "Nome", "Lingua", "Variante", "Rarità", "Tipo carta", "Color",
        "Quantità", "Valore", "Valore totale", "Valore precedente", "Variazione valore", "Variazione %", "Trend prezzo",
        "Fonte prezzo", "Cardmarket idProduct"
    ] if c in candidates.columns]

    sample = candidates[cols].head(max_rows).where(pd.notna(candidates[cols].head(max_rows)), "").to_dict(orient="records")
    summary = {
        "numero_carte_totali_nel_database": int(len(data)),
        "numero_carte_nel_contesto_inviato": int(len(sample)),
        "nota": "I valori economici sono in euro. Rispondi solo usando il contesto fornito; se non basta, dillo chiaramente.",
        "colonne_disponibili": list(data.columns),
        "righe": sample,
    }
    return summary


def ask_openai_about_collection(question, df, token, model):
    context = build_ai_context(df, question)
    system_prompt = (
        "Sei un assistente per una collezione One Piece Card Game. "
        "Rispondi in italiano, in modo pratico e breve. "
        "Usa solo i dati JSON forniti nel messaggio. "
        "I valori sono in euro. Se la domanda richiede un dato non presente, dillo chiaramente."
    )
    user_payload = {
        "domanda": question,
        "contesto_collezione": context,
    }
    response = requests.post(
        "https://api.openai.com/v1/responses",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={
            "model": model or DEFAULT_OPENAI_MODEL,
            "input": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
            ],
            "temperature": 0.1,
        },
        timeout=90,
    )
    if response.status_code >= 400:
        raise RuntimeError(f"Errore API {response.status_code}: {response.text[:1000]}")
    payload = response.json()

    # Responses API: spesso espone output_text, ma tengo fallback difensivi.
    if isinstance(payload, dict) and payload.get("output_text"):
        return str(payload["output_text"]).strip()
    parts = []
    for item in payload.get("output", []) if isinstance(payload, dict) else []:
        for content in item.get("content", []) if isinstance(item, dict) else []:
            if isinstance(content, dict):
                text = content.get("text") or content.get("value")
                if text:
                    parts.append(str(text))
    if parts:
        return "\n".join(parts).strip()
    return json.dumps(payload, ensure_ascii=False, indent=2)[:3000]



st.title("🏴‍☠️ One Piece Card Collection")
st.caption("Dashboard locale per Excel, JSON, valori Cardmarket, trend prezzo, modifiche manuali e log di aggiornamento.")

with st.sidebar:
    st.header("Comandi")
    st.write("Lancia gli script senza aprire PyCharm.")

    if st.button("Crea tutto da zero", use_container_width=True):
        run_script("one_piece_collection_build.py")

    if st.button("Aggiorna solo valori", use_container_width=True):
        run_script("one_piece_collection_update_values.py")

    if st.button("Sync nuove espansioni", use_container_width=True):
        run_script("one_piece_collection_sync.py")

    if st.button("Ricarica dati dashboard", use_container_width=True):
        st.rerun()

    st.divider()
    st.write("Cartelle")
    st.code(
        f"root: {ROOT}\njson: {JSON_DIR}\nout: {Path(OUTPUT_XLSX).parent}\nstg: {STG_DIR}\nlogs: {LOG_DIR}\nbkp: {BKP_DIR}",
        language="text",
    )


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

owned_mask = collection["Quantità"] > 0
owned_collection = collection[owned_mask].copy()

st.info(f"Fonte dati caricata: {source}" + (f" | generato: {generated_at}" if generated_at else ""))
st.caption("Nota generale: i KPI e i grafici di valore/trend considerano le carte possedute, cioè con Quantità > 0, salvo dove è presente un'opzione diversa.")

st.markdown("### Indice dashboard")
st.markdown(
    "Usa le schede qui sotto come indice rapido: **Panoramica**, **Top valore**, **Trend prezzi**, **Gestione carte**, **Domande IA**, **File e log**."
)

tab_overview, tab_top_value, tab_trends, tab_cards, tab_ai, tab_files = st.tabs([
    "📌 Panoramica",
    "💰 Top valore",
    "📈 Trend prezzi",
    "🃏 Gestione carte",
    "🤖 Domande IA",
    "📁 File e log",
])

with tab_overview:
    st.header("Panoramica")
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    cards_total = len(collection)
    cards_owned = len(owned_collection)
    total_qty = int(collection.get("Quantità", 0).sum())
    total_value = float(owned_collection.get("Valore totale", pd.Series(dtype=float)).sum()) if not owned_collection.empty else 0.0
    priced = int((collection.get("Valore", 0) > 0).sum())
    net_delta = float(owned_collection.get("Variazione valore", pd.Series(dtype=float)).sum()) if "Variazione valore" in owned_collection.columns else 0

    k1.metric("Carte", cards_total)
    k2.metric("Carte possedute", cards_owned)
    k3.metric("Quantità totale", total_qty)
    k4.metric("Valore posseduto", euro(total_value))
    k5.metric("Carte con prezzo", priced)
    k6.metric("Delta netto", euro(net_delta))

    if "Trend prezzo" in collection.columns:
        t1, t2, t3, t4 = st.columns(4)
        t1.metric("In aumento", int((owned_collection["Trend prezzo"] == "In aumento").sum()))
        t2.metric("In calo", int((owned_collection["Trend prezzo"] == "In calo").sum()))
        t3.metric("Stabili", int((owned_collection["Trend prezzo"] == "Stabile").sum()))
        t4.metric("Non confrontate", int(owned_collection["Trend prezzo"].isin(["Nuova / non confrontata", "Nessun confronto", "Non posseduta / non confrontata"]).sum()))

    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        if "Rarità" in owned_collection.columns:
            rarity = owned_collection.groupby("Rarità", dropna=False).agg(quantita=("Quantità", "sum")).reset_index()
            rarity = rarity[rarity["Rarità"].astype(str).str.strip() != ""]
            if not rarity.empty:
                fig = px.bar(rarity, x="Rarità", y="quantita", title="Quantità per rarità")
                fig.update_layout(yaxis_title="Quantità", xaxis_title="Rarità", bargap=0.25)
                st.plotly_chart(fig, use_container_width=True)

    with chart_col2:
        if "Lingua" in owned_collection.columns:
            lang = owned_collection.groupby("Lingua", dropna=False).agg(valore=("Valore totale", "sum")).reset_index()
            lang = lang[lang["Lingua"].astype(str).str.strip() != ""]
            if not lang.empty:
                fig = px.pie(lang, names="Lingua", values="valore", title="Valore per lingua (€)")
                st.plotly_chart(fig, use_container_width=True)

    chart_col3, chart_col4 = st.columns(2)
    with chart_col3:
        if "Trend prezzo" in owned_collection.columns:
            trend = owned_collection.groupby("Trend prezzo", dropna=False).size().reset_index(name="carte")
            trend = trend[trend["Trend prezzo"].astype(str).str.strip() != ""]
            if not trend.empty:
                fig = px.bar(trend, x="Trend prezzo", y="carte", title="Distribuzione trend prezzi")
                fig.update_layout(yaxis_title="Carte", xaxis_title="Trend prezzo", bargap=0.25)
                st.plotly_chart(fig, use_container_width=True)

    with chart_col4:
        if "Espansione" in owned_collection.columns:
            exp = owned_collection.groupby("Espansione", dropna=False).agg(valore=("Valore totale", "sum")).reset_index()
            exp = exp.sort_values("valore", ascending=False).head(15)
            if not exp.empty:
                fig = px.bar(exp.sort_values("valore"), x="valore", y="Espansione", orientation="h", title="Top 15 espansioni per valore (€)")
                fig.update_layout(xaxis_title="Valore (€)", yaxis_title="Espansione")
                st.plotly_chart(fig, use_container_width=True)

with tab_top_value:
    st.header("Top carte per valore")
    top_cfg1, top_cfg2, top_cfg3 = st.columns([1, 1, 2])
    with top_cfg1:
        top_x = st.number_input("Numero carte da mostrare", min_value=1, max_value=100, value=10, step=1, key="top_value_x")
    with top_cfg2:
        top_metric = st.selectbox("Tipo valore", ["Valore posseduto (€)", "Valore singola carta (€)"], index=0, key="top_value_metric")
    with top_cfg3:
        include_not_owned_top = st.checkbox("Includi anche carte non possedute", value=False, key="top_value_include_not_owned")

    top_source = collection.copy() if include_not_owned_top else owned_collection.copy()
    value_col = "Valore totale" if top_metric == "Valore posseduto (€)" else "Valore"
    if value_col in top_source.columns and not top_source.empty:
        top_source[value_col] = pd.to_numeric(top_source[value_col], errors="coerce").fillna(0)
        top_cards = top_source[top_source[value_col] > 0].sort_values(value_col, ascending=False).head(int(top_x)).copy()
        if not top_cards.empty:
            label_cols = [c for c in ["ID Carta", "Nome", "Lingua", "Variante"] if c in top_cards.columns]
            top_cards["Carta"] = top_cards[label_cols].astype(str).agg(" | ".join, axis=1) if label_cols else top_cards.index.astype(str)
            m1, m2, m3 = st.columns(3)
            m1.metric("Carta più alta", euro(top_cards[value_col].iloc[0]))
            m2.metric(f"Somma top {int(top_x)}", euro(top_cards[value_col].sum()))
            denom = float((owned_collection["Valore totale"].sum() if "Valore totale" in owned_collection.columns else 0) or 0)
            share = (float(top_cards[value_col].sum()) / denom * 100) if denom > 0 and value_col == "Valore totale" else 0
            m3.metric("Peso sul valore posseduto", f"{share:.1f}%" if value_col == "Valore totale" else "n/d")
            fig = px.bar(
                top_cards.sort_values(value_col, ascending=True),
                x=value_col,
                y="Carta",
                orientation="h",
                title=f"Top {int(top_x)} per {top_metric}",
                text=value_col,
            )
            fig.update_traces(texttemplate="€ %{x:.2f}", textposition="outside", cliponaxis=False)
            fig.update_layout(xaxis_title=top_metric, yaxis_title="Carta", margin=dict(l=10, r=40, t=60, b=10))
            st.plotly_chart(fig, use_container_width=True)
            table_cols = [c for c in ["ID Carta", "Nome", "Lingua", "Variante", "Rarità", "Quantità", "Valore", "Valore totale"] if c in top_cards.columns]
            st.dataframe(top_cards.sort_values(value_col, ascending=False)[table_cols], use_container_width=True, hide_index=True)
        else:
            st.caption("Nessuna carta con valore disponibile per questo filtro.")
    else:
        st.caption("Nessun dato valore disponibile per creare la top carte.")

with tab_trends:
    st.header("Trend prezzi")
    st.caption("Questa sezione considera solo carte possedute, cioè con Quantità > 0.")

    trend_cfg1, trend_cfg2 = st.columns([1, 3])
    with trend_cfg1:
        trend_top_x = st.number_input("Carte da mostrare", min_value=1, max_value=50, value=5, step=1, key="trend_top_x")
    with trend_cfg2:
        st.write("Il numero scelto vale sia per gli aumenti sia per i cali.")

    if "Variazione valore" in owned_collection.columns:
        owned_collection["Variazione valore"] = pd.to_numeric(owned_collection["Variazione valore"], errors="coerce").fillna(0)
        label_cols_base = [c for c in ["ID Carta", "Nome", "Lingua", "Variante"] if c in owned_collection.columns]

        up_col, down_col = st.columns(2)
        with up_col:
            top_up = owned_collection[owned_collection["Variazione valore"] > 0].sort_values("Variazione valore", ascending=False).head(int(trend_top_x)).copy()
            st.subheader(f"Top {int(trend_top_x)} in aumento")
            if not top_up.empty:
                top_up["Carta"] = top_up[label_cols_base].astype(str).agg(" | ".join, axis=1) if label_cols_base else top_up.index.astype(str)
                fig = px.bar(
                    top_up.sort_values("Variazione valore"),
                    x="Variazione valore",
                    y="Carta",
                    orientation="h",
                    title="Aumenti maggiori (€)",
                    text="Variazione valore",
                )
                fig.update_traces(texttemplate="€ %{x:.2f}", textposition="outside", cliponaxis=False)
                fig.update_layout(xaxis_title="Aumento (€)", yaxis_title="Carta", margin=dict(l=10, r=40, t=60, b=10))
                st.plotly_chart(fig, use_container_width=True)
                cols = [c for c in ["ID Carta", "Nome", "Lingua", "Variante", "Valore precedente", "Valore", "Variazione valore", "Variazione %"] if c in top_up.columns]
                st.dataframe(top_up[cols], use_container_width=True, hide_index=True)
            else:
                st.caption("Nessun aumento disponibile.")

        with down_col:
            top_down = owned_collection[owned_collection["Variazione valore"] < 0].sort_values("Variazione valore", ascending=True).head(int(trend_top_x)).copy()
            st.subheader(f"Top {int(trend_top_x)} in calo")
            if not top_down.empty:
                top_down["Carta"] = top_down[label_cols_base].astype(str).agg(" | ".join, axis=1) if label_cols_base else top_down.index.astype(str)
                top_down["Calo (€)"] = top_down["Variazione valore"].abs()
                fig = px.bar(
                    top_down.sort_values("Calo (€)"),
                    x="Calo (€)",
                    y="Carta",
                    orientation="h",
                    title="Cali maggiori (€)",
                    text="Calo (€)",
                )
                fig.update_traces(texttemplate="€ %{x:.2f}", textposition="outside", cliponaxis=False)
                fig.update_layout(xaxis_title="Calo (€)", yaxis_title="Carta", margin=dict(l=10, r=40, t=60, b=10))
                st.plotly_chart(fig, use_container_width=True)
                cols = [c for c in ["ID Carta", "Nome", "Lingua", "Variante", "Valore precedente", "Valore", "Variazione valore", "Variazione %"] if c in top_down.columns]
                st.dataframe(top_down[cols], use_container_width=True, hide_index=True)
            else:
                st.caption("Nessun calo disponibile.")
    else:
        st.caption("Nessun confronto disponibile.")

    st.subheader("Valore collezione nel tempo")
    history_path = Path(VALUE_HISTORY_CSV)
    if history_path.exists():
        try:
            history = pd.read_csv(history_path, encoding="utf-8-sig")
            if not history.empty and "Data" in history.columns and "Valore collezione (€)" in history.columns:
                history["Data"] = pd.to_datetime(history["Data"], errors="coerce")
                history["Valore collezione (€)"] = pd.to_numeric(history["Valore collezione (€)"], errors="coerce").fillna(0)
                history = history.dropna(subset=["Data"]).sort_values("Data")
                fig = px.line(history, x="Data", y="Valore collezione (€)", markers=True, title="Valore collezione nel tempo (€)")
                fig.update_layout(yaxis_title="Valore (€)", xaxis_title="Data")
                st.plotly_chart(fig, use_container_width=True)
                st.dataframe(history.tail(10), use_container_width=True, hide_index=True)
            else:
                st.caption("Storico presente ma ancora senza dati utili.")
        except Exception as exc:
            st.warning(f"Non riesco a leggere lo storico valori: {exc}")
    else:
        st.caption("Lo storico viene creato dal prossimo build/update/sync o salvataggio manuale.")

with tab_cards:
    st.header("Gestione carte")

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

    preferred_cols = [
        "ID Carta", "Espansione", "Numero", "Nome", "Lingua", "Variante", "Rarità", "Tipo carta", "Color", "Quantità",
        "Valore", "Valore totale", "Valore precedente", "Variazione valore", "Variazione %", "Trend prezzo",
        "Rarità JP ufficiale", "Rarità JP candidate", "Fonte rarità JP",
    ]
    show_cols = [c for c in preferred_cols if c in filtered.columns]
    show_cols += [c for c in filtered.columns if c not in show_cols]

    st.caption("Puoi modificare tutti i campi visibili. I valori economici sono in euro (€). Dopo la modifica premi Salva modifiche.")

    def _safe_text(value):
        if pd.isna(value):
            return ""
        return str(value)

    def _card_label(idx, row):
        parts = []
        for col in ["ID Carta", "Nome", "Lingua", "Variante", "Rarità"]:
            if col in row.index:
                val = _safe_text(row[col]).strip()
                if val:
                    parts.append(val)
        label = " | ".join(parts) if parts else f"Carta senza nome #{idx}"
        return f"#{idx} - {label}"

    def _guess_number_from_card_id(card_id):
        text = str(card_id or "").strip().upper().replace(".", "-").replace("_", "-")
        if "-" not in text:
            return ""
        tail = text.split("-")[-1]
        return tail.zfill(3) if tail.isdigit() else tail

    def _guess_expansion_from_card_id(card_id):
        text = str(card_id or "").strip().upper().replace(".", "-").replace("_", "-")
        if "-" not in text:
            return ""
        return text.split("-")[0]

    manage_tab_edit, manage_tab_add_delete = st.tabs(["Modifica tabella", "Aggiungi / elimina"])

    with manage_tab_add_delete:
        st.caption("Usa questi comandi per aggiungere una carta manuale o rimuovere una carta specifica dalla collezione. Il salvataggio crea sempre un backup automatico.")
        add_tab, del_tab = st.tabs(["Aggiungi nuova carta", "Elimina carta"])

        with add_tab:
            with st.form("add_card_form", clear_on_submit=True):
                c1, c2, c3, c4 = st.columns(4)
                new_id = c1.text_input("ID Carta", placeholder="OP16-042")
                new_name = c2.text_input("Nome", placeholder="Prisoner of Impel Down")
                new_lang = c3.selectbox("Lingua", ["EN", "JP", ""], index=0)
                new_variant = c4.text_input("Variante", value="Base")

                c5, c6, c7, c8 = st.columns(4)
                new_exp = c5.text_input("Espansione", value=_guess_expansion_from_card_id(new_id))
                new_num = c6.text_input("Numero", value=_guess_number_from_card_id(new_id))
                new_rarity = c7.text_input("Rarità", placeholder="C, UC, R, SR, SEC, TR...")
                new_type = c8.text_input("Tipo carta", placeholder="CHARACTER, EVENT, LEADER...")

                c9, c10, c11, c12 = st.columns(4)
                new_color = c9.text_input("Color", placeholder="Blue, Red, Yellow...")
                new_qty = c10.number_input("Quantità", min_value=0, step=1, value=1)
                new_value = c11.number_input("Valore (€)", min_value=0.0, step=0.01, value=0.0, format="%.2f")
                new_source = c12.text_input("Fonte prezzo", value="Manuale")

                submitted_add = st.form_submit_button("Aggiungi carta", type="primary", use_container_width=True)

            if submitted_add:
                try:
                    if not str(new_id).strip() and not str(new_name).strip():
                        st.error("Inserisci almeno ID Carta o Nome.")
                    else:
                        updated = collection.copy()
                        required_cols = [
                            "ID Carta", "Espansione", "Numero", "Nome", "Lingua", "Variante", "Rarità", "Tipo carta", "Color",
                            "Quantità", "Valore", "Valore totale", "Fonte prezzo"
                        ]
                        for col in required_cols:
                            if col not in updated.columns:
                                updated[col] = "" if col not in {"Quantità", "Valore", "Valore totale"} else 0

                        card_id = str(new_id).strip().upper().replace(".", "-").replace("_", "-")
                        row = {col: "" for col in updated.columns}
                        row.update({
                            "ID Carta": card_id,
                            "Espansione": str(new_exp).strip() or _guess_expansion_from_card_id(card_id),
                            "Numero": str(new_num).strip() or _guess_number_from_card_id(card_id),
                            "Nome": str(new_name).strip(),
                            "Lingua": str(new_lang).strip(),
                            "Variante": str(new_variant).strip(),
                            "Rarità": str(new_rarity).strip(),
                            "Tipo carta": str(new_type).strip(),
                            "Color": str(new_color).strip(),
                            "Quantità": int(new_qty),
                            "Valore": float(new_value),
                            "Valore totale": int(new_qty) * float(new_value),
                            "Fonte prezzo": str(new_source).strip() or "Manuale",
                        })
                        updated = pd.concat([updated, pd.DataFrame([row])], ignore_index=True)
                        save_collection_from_streamlit(updated, source="streamlit_add_card")
                        st.success(f"Carta aggiunta: {row.get('ID Carta', '')} {row.get('Nome', '')}".strip())
                        st.cache_data.clear()
                        time.sleep(0.5)
                        st.rerun()
                except Exception as exc:
                    st.error(f"Errore durante l'aggiunta carta: {exc}")

        with del_tab:
            if collection.empty:
                st.caption("Nessuna carta da eliminare.")
            else:
                delete_options = [(int(idx), _card_label(idx, row)) for idx, row in collection.iterrows()]
                selected_delete = st.selectbox(
                    "Carta da eliminare",
                    options=delete_options,
                    format_func=lambda item: item[1],
                    index=0,
                    help="L'eliminazione rimuove la carta dal JSON e dall'Excel finale dopo backup automatico.",
                )
                confirm_delete = st.checkbox("Confermo di voler eliminare questa carta", value=False)
                if st.button("Elimina carta selezionata", type="secondary", use_container_width=True, disabled=not confirm_delete):
                    try:
                        row_id = int(selected_delete[0])
                        deleted_label = selected_delete[1]
                        updated = collection.drop(index=row_id).reset_index(drop=True)
                        save_collection_from_streamlit(updated, source="streamlit_delete_card")
                        st.success(f"Carta eliminata: {deleted_label}")
                        st.cache_data.clear()
                        time.sleep(0.5)
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Errore durante l'eliminazione carta: {exc}")

    with manage_tab_edit:
        editor_df = filtered[show_cols].copy()
        editor_df.insert(0, "_row_id", filtered.index.astype(int))

        column_config = {
            "_row_id": None,
            "Quantità": st.column_config.NumberColumn("Quantità", min_value=0, step=1, format="%d", help="Quante copie possiedi."),
            "Valore": money_column("Valore (€)"),
            "Valore totale": money_column("Valore totale (€)"),
            "Valore precedente": money_column("Valore precedente (€)"),
            "Variazione valore": money_column("Variazione valore (€)"),
            "Variazione %": st.column_config.NumberColumn("Variazione %", format="%.2f%%"),
        }

        edited_df = st.data_editor(
            editor_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            disabled=["_row_id"],
            column_config=column_config,
            key="cards_full_editor",
        )

        save_col, hint_col = st.columns([1, 3])
        with save_col:
            save_pressed = st.button("Salva modifiche", type="primary", use_container_width=True)
        with hint_col:
            st.caption("Il salvataggio aggiorna out/one_piece_collection.json e out/one_piece_collection.xlsx, con backup automatico prima della modifica. Valore totale viene ricalcolato da Quantità × Valore.")

        if save_pressed:
            try:
                updated = collection.copy()
                changes = 0
                numeric_cols = {"Quantità", "Valore", "Valore totale", "Valore precedente", "Variazione valore", "Variazione %", "CM_Low", "CM_Trend", "CM_Avg", "CM_Avg1", "CM_Avg7", "CM_Avg30", "Power", "Counter", "Cost", "Life"}
                for _, row in edited_df.iterrows():
                    row_id = int(row["_row_id"])
                    for col in show_cols:
                        if col not in updated.columns or col == "_row_id":
                            continue
                        new_value = row.get(col, "")
                        old_value = updated.at[row_id, col]
                        if col in numeric_cols:
                            new_num = pd.to_numeric(new_value, errors="coerce")
                            old_num = pd.to_numeric(old_value, errors="coerce")
                            if pd.isna(new_num):
                                new_value = 0
                            else:
                                new_value = int(new_num) if col == "Quantità" else float(new_num)
                            old_cmp = 0 if pd.isna(old_num) else (int(old_num) if col == "Quantità" else float(old_num))
                            changed = new_value != old_cmp
                        else:
                            new_value = "" if pd.isna(new_value) else str(new_value)
                            old_cmp = "" if pd.isna(old_value) else str(old_value)
                            changed = new_value != old_cmp
                        if changed:
                            updated.at[row_id, col] = new_value
                            changes += 1

                if "Quantità" in updated.columns and "Valore" in updated.columns:
                    updated["Quantità"] = pd.to_numeric(updated["Quantità"], errors="coerce").fillna(0).clip(lower=0).round(0).astype(int)
                    updated["Valore"] = pd.to_numeric(updated["Valore"], errors="coerce").fillna(0)
                    updated["Valore totale"] = updated["Quantità"] * updated["Valore"]
                save_collection_from_streamlit(updated)
                st.success(f"Modifiche salvate. Campi modificati: {changes}. Excel e JSON aggiornati.")
                st.cache_data.clear()
                time.sleep(0.5)
                st.rerun()
            except Exception as exc:
                st.error(f"Errore durante il salvataggio modifiche: {exc}")

with tab_ai:
    st.header("Domande sulla collezione con IA")
    st.caption("Se configuri un token OpenAI, puoi fare domande testuali sui dati caricati. Il token viene salvato localmente in config/openai_api_key.txt, file escluso da Git.")
    with st.expander("Configura IA e fai una domanda", expanded=True):
        saved_token = load_saved_token()
        saved_model = load_saved_model()
        masked = "Token salvato" if saved_token else "Nessun token salvato"
        st.info(masked)

        c_ai1, c_ai2 = st.columns([2, 1])
        with c_ai1:
            token_input = st.text_input("OpenAI API token", value="", type="password", placeholder="sk-...", help="Lascia vuoto per non modificare il token salvato.")
        with c_ai2:
            model_input = st.text_input("Modello", value=saved_model or DEFAULT_OPENAI_MODEL)

        b_ai1, b_ai2 = st.columns([1, 1])
        with b_ai1:
            if st.button("Salva token/modello", use_container_width=True):
                try:
                    if token_input.strip():
                        save_token(token_input)
                    save_model(model_input)
                    st.success("Configurazione IA salvata in config/.")
                except Exception as exc:
                    st.error(f"Errore salvataggio configurazione IA: {exc}")
        with b_ai2:
            if st.button("Rimuovi token salvato", use_container_width=True):
                try:
                    save_token("")
                    st.success("Token rimosso.")
                except Exception as exc:
                    st.error(f"Errore rimozione token: {exc}")

        question = st.text_area("Domanda", placeholder="Esempio: qual è la carta più costosa del set OP03?", height=90)
        if st.button("Chiedi all'IA", type="primary", use_container_width=True):
            token = token_input.strip() or load_saved_token()
            model = (model_input or load_saved_model() or DEFAULT_OPENAI_MODEL).strip()
            if not token:
                st.warning("Sezione disponibile solo impostando un token OpenAI. Inseriscilo qui sopra e premi Salva token/modello.")
            elif not question.strip():
                st.warning("Scrivi una domanda prima di inviare.")
            else:
                with st.spinner("Interrogo l'IA sui dati della collezione..."):
                    try:
                        answer = ask_openai_about_collection(question, collection, token, model)
                        st.markdown(answer)
                    except Exception as exc:
                        msg = str(exc)
                        if "insufficient_quota" in msg or "quota" in msg.lower():
                            st.error("Quota API OpenAI esaurita o billing API non configurato. La dashboard funziona comunque senza IA.")
                        else:
                            st.error(f"Errore chiamata IA: {exc}")

with tab_files:
    st.header("File e log")
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
        else:
            st.caption("Excel finale non trovato.")

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
        else:
            st.caption("JSON finale non trovato.")

    with log_col:
        logs = latest_log_files()
        if logs:
            selected = st.selectbox("Ultimi log", logs, format_func=lambda p: p.name)
            with st.expander("Mostra log selezionato", expanded=True):
                st.code(selected.read_text(encoding="utf-8", errors="replace")[-20000:], language="text")
        else:
            st.caption("Nessun log trovato.")
