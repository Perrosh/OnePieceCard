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



# =============================================================================
# Interprete locale domande database: funziona senza servizi esterni.
# =============================================================================

def _qnorm(text):
    text = str(text or "").lower()
    repl = {"à": "a", "è": "e", "é": "e", "ì": "i", "ò": "o", "ù": "u", "’": "'", "–": "-", "—": "-"}
    for a, b in repl.items():
        text = text.replace(a, b)
    return text


def _extract_intent_number(q, default=10):
    import re
    qn = _qnorm(q)
    m = re.search(r"\btop\s*(\d{1,3})\b", qn)
    if m:
        return max(1, min(200, int(m.group(1))))
    m = re.search(r"\b(prime|primi|mostra|vedere|lista|elenca)\s+(\d{1,3})\b", qn)
    if m:
        return max(1, min(200, int(m.group(2))))
    return default


def _extract_expansion_patterns(q):
    import re
    qn = _qnorm(q)
    patterns = set()
    for prefix, number in re.findall(r"\b(OP|ST|EB|PRB)[-\s]?(\d{1,2})\b", qn, flags=re.IGNORECASE):
        prefix = prefix.upper()
        n = int(number)
        patterns.add(f"{prefix}{n:02d}".lower())
        patterns.add(f"{prefix}-{n:02d}".lower())
    for left, right in re.findall(r"\b(OP\d{1,2})[-\s]?(EB\d{1,2})\b", qn, flags=re.IGNORECASE):
        patterns.add(f"{left.upper()}-{right.upper()}".lower())
        patterns.add(f"{left.upper()}{right.upper()}".lower())
    return patterns


def _filter_database_from_question(df, q):
    import re
    data = df.copy()
    qn = _qnorm(q)

    wants_owned = any(w in qn for w in ["possed", "che ho", "mie carte", "mia collezione", "collezione", "quantita > 0"])
    value_or_trend_question = any(w in qn for w in ["costos", "val", "prezz", "aument", "cal", "trend", "top"])
    if (wants_owned or value_or_trend_question) and "Quantità" in data.columns:
        data = data[pd.to_numeric(data["Quantità"], errors="coerce").fillna(0) > 0].copy()

    patterns = _extract_expansion_patterns(qn)
    if patterns:
        mask = pd.Series(False, index=data.index)
        for col in ["ID Carta", "Espansione", "Set", "Codice", "Nome"]:
            if col in data.columns:
                text = data[col].astype(str).str.lower()
                text_compact = text.str.replace("-", "", regex=False).str.replace(" ", "", regex=False)
                for pat in patterns:
                    mask = mask | text.str.contains(pat, na=False)
                    mask = mask | text_compact.str.contains(pat.replace("-", "").replace(" ", ""), na=False)
        data = data[mask].copy()

    if re.search(r"\bjp\b|giappon", qn) and "Lingua" in data.columns:
        data = data[data["Lingua"].astype(str).str.upper().eq("JP")].copy()
    if re.search(r"\ben\b|ingles", qn) and "Lingua" in data.columns:
        data = data[data["Lingua"].astype(str).str.upper().eq("EN")].copy()

    rarity_tokens = ["SEC", "SR", "R", "UC", "C", "L", "TR", "SP", "P", "DON"]
    if "Rarità" in data.columns:
        for rar in rarity_tokens:
            if re.search(rf"\b{re.escape(rar.lower())}\b", qn):
                data = data[data["Rarità"].astype(str).str.upper().eq(rar)].copy()
                break

    colors = {"red": ["red", "rosso", "rossa"], "green": ["green", "verde"], "blue": ["blue", "blu"], "purple": ["purple", "viola"], "black": ["black", "nero", "nera"], "yellow": ["yellow", "giallo", "gialla"]}
    color_col = "Color" if "Color" in data.columns else ("Colore" if "Colore" in data.columns else None)
    if color_col:
        for canonical, words in colors.items():
            if any(w in qn for w in words):
                data = data[data[color_col].astype(str).str.lower().str.contains(canonical, na=False)].copy()
                break
    return data


def _answer_with_table(title, df, columns=None, max_rows=10):
    if columns is None:
        columns = [c for c in ["ID Carta", "Espansione", "Numero", "Nome", "Lingua", "Variante", "Rarità", "Quantità", "Valore", "Valore totale", "Variazione valore", "Trend prezzo"] if c in df.columns]
    shown = df[columns].head(max_rows).copy() if columns else df.head(max_rows).copy()
    rename = {"Valore": "Valore (€)", "Valore totale": "Valore totale (€)", "Variazione valore": "Variazione valore (€)"}
    shown = shown.rename(columns={k: v for k, v in rename.items() if k in shown.columns})
    return title, shown


def answer_local_database_question(question, df):
    import re
    q = str(question or "").strip()
    qn = _qnorm(q)
    if not q:
        return False, "", None
    data = df.copy()
    for col in ["Quantità", "Valore", "Valore totale", "Valore precedente", "Variazione valore", "Variazione %"]:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)
    if "Valore totale" not in data.columns and {"Quantità", "Valore"}.issubset(data.columns):
        data["Valore totale"] = data["Quantità"] * data["Valore"]
    filtered = _filter_database_from_question(data, q)
    n = _extract_intent_number(q, default=10)
    if filtered.empty:
        return True, "Non ho trovato carte che rispettano i filtri della domanda.", None

    if any(w in qn for w in ["quante", "numero", "conteggio"]):
        owned = int((pd.to_numeric(filtered.get("Quantità", 0), errors="coerce").fillna(0) > 0).sum()) if "Quantità" in filtered.columns else 0
        qty = int(pd.to_numeric(filtered.get("Quantità", 0), errors="coerce").fillna(0).sum()) if "Quantità" in filtered.columns else 0
        total = float(pd.to_numeric(filtered.get("Valore totale", 0), errors="coerce").fillna(0).sum()) if "Valore totale" in filtered.columns else 0.0
        msg = f"Ho trovato **{len(filtered)} carte** nel filtro richiesto. Carte possedute: **{owned}**. Quantità totale: **{qty}**. Valore posseduto: **€ {total:,.2f}**.".replace(",", "X").replace(".", ",").replace("X", ".")
        return True, msg, None

    if any(w in qn for w in ["valore totale", "quanto vale", "totale collezione", "valore collezione"]):
        total = float(pd.to_numeric(filtered.get("Valore totale", 0), errors="coerce").fillna(0).sum()) if "Valore totale" in filtered.columns else 0.0
        msg = f"Il valore totale delle carte nel filtro richiesto è **€ {total:,.2f}**.".replace(",", "X").replace(".", ",").replace("X", ".")
        return True, msg, None

    if any(w in qn for w in ["aument", "salit", "cresciut"]):
        if "Variazione valore" not in filtered.columns:
            return True, "Nel database non trovo la colonna `Variazione valore` per calcolare gli aumenti.", None
        out = filtered[pd.to_numeric(filtered["Variazione valore"], errors="coerce").fillna(0) > 0].sort_values("Variazione valore", ascending=False)
        title, table = _answer_with_table(f"Top {min(n, len(out))} carte in aumento", out, max_rows=n)
        return True, title, table

    if any(w in qn for w in ["cal", "sces", "diminuit", "ribass"]):
        if "Variazione valore" not in filtered.columns:
            return True, "Nel database non trovo la colonna `Variazione valore` per calcolare i cali.", None
        out = filtered[pd.to_numeric(filtered["Variazione valore"], errors="coerce").fillna(0) < 0].sort_values("Variazione valore", ascending=True)
        title, table = _answer_with_table(f"Top {min(n, len(out))} carte in calo", out, max_rows=n)
        return True, title, table

    if any(w in qn for w in ["costos", "car", "valgono", "vale di piu", "piu valore", "top"]):
        use_total = any(w in qn for w in ["valore totale", "valore posseduto", "quantita", "complessivo", "totale"])
        sort_col = "Valore totale" if use_total and "Valore totale" in filtered.columns else "Valore"
        if sort_col not in filtered.columns:
            return True, f"Nel database non trovo la colonna `{sort_col}`.", None
        out = filtered.sort_values(sort_col, ascending=False)
        if any(w in qn for w in ["qual e", "quale", "piu costosa", "piu cara", "massima", "massimo"]) and not re.search(r"\btop\s*\d+", qn):
            top = out.head(1)
            row = top.iloc[0]
            name = row.get("Nome", "Carta")
            cid = row.get("ID Carta", "")
            exp = row.get("Espansione", "")
            lang = row.get("Lingua", "")
            var = row.get("Variante", "")
            qty = row.get("Quantità", "")
            val = float(row.get(sort_col, 0) or 0)
            label = "valore posseduto" if sort_col == "Valore totale" else "valore unitario"
            msg = f"La carta più alta nel filtro richiesto è **{name}**"
            details = [str(x) for x in [cid, exp, lang, var] if str(x or "").strip()]
            if details:
                msg += " (" + " | ".join(details) + ")"
            msg += f", con {label} **€ {val:,.2f}**".replace(",", "X").replace(".", ",").replace("X", ".")
            if qty != "":
                msg += f". Quantità: **{qty}**."
            return True, msg, _answer_with_table("Dettaglio", top, max_rows=1)[1]
        title, table = _answer_with_table(f"Top {min(n, len(out))} carte per {sort_col.lower()} (€)", out, max_rows=n)
        return True, title, table

    if any(w in qn for w in ["possed", "che ho", "lista", "elenca", "mostra"]):
        sort_col = "Valore totale" if "Valore totale" in filtered.columns else filtered.columns[0]
        out = filtered.sort_values(sort_col, ascending=False)
        title, table = _answer_with_table(f"Prime {min(n, len(out))} carte trovate", out, max_rows=n)
        return True, title, table

    examples = """Posso rispondere localmente a domande tipo:

- `qual è la carta più costosa dell'espansione OP03 tra le carte che possiedo?`
- `top 10 carte OP16 per valore posseduto`
- `quali carte JP possiedo che valgono di più?`
- `quante carte SEC possiedo?`
- `top 5 carte in aumento`
- `valore totale OP03`

Le domande non riconosciute mostrano questi esempi: il motore è locale e basato sui campi del database."""
    return True, examples, None


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
    "Usa le schede qui sotto come indice rapido: **Panoramica**, **Top valore**, **Trend prezzi**, **Gestione carte**, **Domande database**, **File e log**."
)

tab_overview, tab_top_value, tab_trends, tab_cards, tab_ai, tab_files = st.tabs([
    "📌 Panoramica",
    "💰 Top valore",
    "📈 Trend prezzi",
    "🃏 Gestione carte",
    "🔎 Domande database",
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
    st.header("Domande sul database")
    st.caption("Questa sezione funziona in locale: interpreta la domanda con regole e Pandas, senza chiamate esterne.")

    question = st.text_area(
        "Domanda",
        placeholder="Esempio: qual è la carta più costosa dell'espansione OP03 tra le carte che possiedo?",
        height=90,
    )

    if st.button("Rispondi dal database", type="primary", use_container_width=True):
        if not question.strip():
            st.warning("Scrivi una domanda prima di inviare.")
        else:
            handled, answer, answer_df = answer_local_database_question(question, collection)
            st.markdown(answer)
            if answer_df is not None and not answer_df.empty:
                st.dataframe(answer_df, use_container_width=True, hide_index=True)

    with st.expander("Esempi di domande supportate"):
        st.markdown(
            """
            - qual è la carta più costosa dell'espansione OP03 tra le carte che possiedo?
            - top 10 carte OP16 per valore posseduto
            - quali carte JP possiedo che valgono di più?
            - quante carte SEC possiedo?
            - top 5 carte in aumento
            - valore totale OP03
            """
        )

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
