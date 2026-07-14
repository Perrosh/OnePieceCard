import json
import os
import re
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
    find_json,
    PRICE_GUIDE_PATTERN,
    PRODUCTS_SINGLES_PATTERN,
    normalize_code,
    normalize_idproduct_for_price_key,
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

THEME_PALETTES = {
    "Predefinito": {
        "app": "linear-gradient(135deg, #f8fafc 0%, #eef2ff 45%, #f8fafc 100%)",
        "sidebar": "linear-gradient(180deg, #ffffff 0%, #f1f5f9 100%)",
        "text": "#0f172a",
        "muted": "#475569",
        "card": "rgba(255,255,255,0.82)",
        "border": "rgba(15,23,42,0.14)",
        "tab": "rgba(15,23,42,0.06)",
        "tab_selected": "rgba(37,99,235,0.16)",
        "accent": "#2563eb",
        "warning_bg": "rgba(245,158,11,0.16)",
        "ok_bg": "rgba(34,197,94,0.12)",
    },
    "Chiaro": {
        "app": "linear-gradient(135deg, #fff7ed 0%, #fefce8 45%, #eff6ff 100%)",
        "sidebar": "linear-gradient(180deg, #fff7ed 0%, #ffffff 100%)",
        "text": "#1f2937",
        "muted": "#64748b",
        "card": "rgba(255,255,255,0.88)",
        "border": "rgba(31,41,55,0.13)",
        "tab": "rgba(251,146,60,0.10)",
        "tab_selected": "rgba(251,146,60,0.24)",
        "accent": "#ea580c",
        "warning_bg": "rgba(245,158,11,0.16)",
        "ok_bg": "rgba(34,197,94,0.12)",
    },
    "Scuro": {
        "app": "radial-gradient(circle at top left, #243447 0, #101722 38%, #0b1018 100%)",
        "sidebar": "linear-gradient(180deg, #111827 0%, #0b1220 100%)",
        "text": "#edf2f7",
        "muted": "#cbd5e1",
        "card": "rgba(255,255,255,0.07)",
        "border": "rgba(255,255,255,0.12)",
        "tab": "rgba(255,255,255,0.06)",
        "tab_selected": "rgba(255,255,255,0.18)",
        "accent": "#38bdf8",
        "warning_bg": "rgba(245,158,11,0.18)",
        "ok_bg": "rgba(34,197,94,0.14)",
    },
    "Mare": {
        "app": "radial-gradient(circle at top left, #0f766e 0%, #083344 42%, #020617 100%)",
        "sidebar": "linear-gradient(180deg, #042f2e 0%, #020617 100%)",
        "text": "#ecfeff",
        "muted": "#a7f3d0",
        "card": "rgba(20,184,166,0.13)",
        "border": "rgba(125,211,252,0.26)",
        "tab": "rgba(20,184,166,0.13)",
        "tab_selected": "rgba(45,212,191,0.28)",
        "accent": "#2dd4bf",
        "warning_bg": "rgba(251,191,36,0.18)",
        "ok_bg": "rgba(45,212,191,0.16)",
    },
    "Wanted": {
        "app": "linear-gradient(135deg, #3b2517 0%, #7c4a22 45%, #f4d19b 100%)",
        "sidebar": "linear-gradient(180deg, #2b1a10 0%, #5b3419 100%)",
        "text": "#fff7ed",
        "muted": "#fed7aa",
        "card": "rgba(120,53,15,0.28)",
        "border": "rgba(254,215,170,0.34)",
        "tab": "rgba(254,215,170,0.15)",
        "tab_selected": "rgba(251,191,36,0.28)",
        "accent": "#fbbf24",
        "warning_bg": "rgba(251,191,36,0.22)",
        "ok_bg": "rgba(34,197,94,0.14)",
    },
    "One Piece": {
        "app": "radial-gradient(circle at top left, #0ea5e9 0%, #075985 26%, #111827 58%, #3b1d0f 100%)",
        "sidebar": "linear-gradient(180deg, #0f172a 0%, #172554 46%, #451a03 100%)",
        "text": "#fff7ed",
        "muted": "#fde68a",
        "card": "rgba(15, 23, 42, 0.58)",
        "border": "rgba(251, 191, 36, 0.42)",
        "tab": "rgba(14, 165, 233, 0.14)",
        "tab_selected": "rgba(251, 191, 36, 0.24)",
        "accent": "#facc15",
        "warning_bg": "rgba(251, 191, 36, 0.20)",
        "ok_bg": "rgba(34,197,94,0.16)",
    },
}

if "dashboard_theme" not in st.session_state:
    st.session_state["dashboard_theme"] = "One Piece"

_header_left, _header_right = st.columns([6, 1.45], vertical_alignment="center")
with _header_left:
    st.title("🏴‍☠️ One Piece Card Collection")
with _header_right:
    selected_theme = st.selectbox(
        "Tema",
        list(THEME_PALETTES.keys()),
        index=list(THEME_PALETTES.keys()).index(st.session_state.get("dashboard_theme", "One Piece")) if st.session_state.get("dashboard_theme", "One Piece") in THEME_PALETTES else list(THEME_PALETTES.keys()).index("One Piece"),
        key="dashboard_theme",
        label_visibility="collapsed",
    )

pal = THEME_PALETTES[selected_theme]
st.markdown(
    f"""
    <style>
    .stApp {{
        background: {pal['app']};
        color: {pal['text']};
    }}
    section[data-testid="stSidebar"] {{
        display: none !important;
    }}
    div[data-testid="collapsedControl"] {{
        display: none !important;
    }}
    .block-container {{
        padding-top: 1.3rem;
    }}
    h1, h2, h3, h4, h5, h6, p, label, span {{
        color: {pal['text']};
    }}
    div[data-testid="stMetric"] {{
        background: {pal['card']};
        border: 1px solid {pal['border']};
        border-radius: 16px;
        padding: 14px 16px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.18);
    }}
    div[data-testid="stMetric"] label, div[data-testid="stMetric"] p {{
        color: {pal['muted']} !important;
    }}
    .stTabs [data-baseweb="tab-list"] {{
        gap: 8px;
        background: {pal['card']};
        border: 1px solid {pal['border']};
        border-radius: 14px;
        padding: 8px;
    }}
    .stTabs [data-baseweb="tab"] {{
        background: {pal['tab']};
        border-radius: 12px;
        padding: 8px 14px;
    }}
    .stTabs [aria-selected="true"] {{
        background: {pal['tab_selected']};
        box-shadow: inset 0 -2px 0 {pal['accent']};
    }}
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {{
        border-radius: 14px;
        overflow: hidden;
        border: 1px solid {pal['border']};
    }}
    .op-card {{
        background: {pal['card']};
        border: 1px solid {pal['border']};
        border-radius: 16px;
        padding: 14px 16px;
        margin: 8px 0 14px 0;
    }}
    .op-warning {{
        background: {pal['warning_bg']};
        border: 1px solid rgba(245, 158, 11, 0.45);
        border-radius: 14px;
        padding: 12px 14px;
        margin: 8px 0 12px 0;
    }}
    .op-ok {{
        background: {pal['ok_bg']};
        border: 1px solid rgba(34, 197, 94, 0.35);
        border-radius: 14px;
        padding: 12px 14px;
        margin: 8px 0 12px 0;
    }}
    div[data-testid="stVerticalBlock"]:has(.op-toolbar-marker) {{
        position: sticky;
        top: 0;
        z-index: 999;
        background: {pal['card']};
        border: 1px solid {pal['border']};
        border-radius: 0 0 18px 18px;
        padding: 10px 12px 12px 12px;
        margin-bottom: 12px;
        backdrop-filter: blur(14px);
        box-shadow: 0 12px 28px rgba(0,0,0,0.26);
    }}
    .op-toolbar-marker {{
        height: 0;
        margin: 0;
        padding: 0;
    }}
    .op-run-summary {{
        background: {pal['ok_bg']};
        border: 1px solid rgba(34, 197, 94, 0.36);
        border-radius: 14px;
        padding: 12px 14px;
        margin: 8px 0 12px 0;
    }}
    button[kind="primary"] {{
        border: 1px solid {pal['accent']};
    }}
    </style>
    """,
    unsafe_allow_html=True,
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




def _extract_card_id_from_cardmarket_name(name):
    """Estrae OPxx-xxx/STxx-xxx/EBxx-xxx/PRBxx-xxx/P-xxx dal nome prodotto Cardmarket."""
    if not isinstance(name, str):
        return ""
    matches = re.findall(r"\(([A-Z]{1,5}-?\d{2}-\d{3}|P-?\d{3})\)", name)
    return normalize_code(matches[-1]) if matches else ""


def _clean_cardmarket_name(name):
    if not isinstance(name, str):
        return ""
    return re.sub(r"\s*\(([A-Z]{1,5}-?\d{2}-\d{3}|P-?\d{3})\)\s*$", "", name).strip()


def _load_json_file(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def debug_cardmarket_update_for_idproduct(idproduct, collection_df):
    """Simula in modo leggibile la logica di Aggiorna valori per un idProduct.

    Non modifica nulla: legge i JSON correnti, cerca il prodotto, cerca il prezzo,
    trova eventuali righe nel database e mostra quali soli campi di mercato verrebbero aggiornati.
    """
    result = {
        "ok": False,
        "messages": [],
        "files": {},
        "product": {},
        "price": {},
        "derived": {},
        "rows_by_idproduct": pd.DataFrame(),
        "rows_by_card_id": pd.DataFrame(),
        "updates_preview": pd.DataFrame(),
    }

    normalized = normalize_idproduct_for_price_key(idproduct)
    result["derived"]["Cardmarket idProduct richiesto"] = str(idproduct).strip()
    result["derived"]["Cardmarket idProduct normalizzato"] = normalized

    if not normalized:
        result["messages"].append("Inserisci un Cardmarket idProduct valido.")
        return result

    try:
        products_file = find_json(PRODUCTS_SINGLES_PATTERN)
        price_file = find_json(PRICE_GUIDE_PATTERN)
        result["files"] = {
            "Products singles JSON": products_file,
            "Price guide JSON": price_file,
        }
    except Exception as e:
        result["messages"].append(f"JSON Cardmarket mancanti o non leggibili: {e}")
        return result

    try:
        products_data = _load_json_file(products_file)
        price_data = _load_json_file(price_file)
    except Exception as e:
        result["messages"].append(f"Errore lettura JSON: {e}")
        return result

    products = products_data.get("products", []) if isinstance(products_data, dict) else []
    price_guides = price_data.get("priceGuides", []) if isinstance(price_data, dict) else []

    product = None
    for item in products:
        if normalize_idproduct_for_price_key(item.get("idProduct", "")) == normalized:
            product = item
            break

    price = None
    for item in price_guides:
        if normalize_idproduct_for_price_key(item.get("idProduct", "")) == normalized:
            price = item
            break

    if not product:
        result["messages"].append("Prodotto non trovato in products_singles JSON.")
    if not price:
        result["messages"].append("Prezzo non trovato in price_guide JSON.")

    if product:
        product_name = product.get("name", "")
        card_id = _extract_card_id_from_cardmarket_name(product_name)
        clean_name = _clean_cardmarket_name(product_name)
        result["product"] = {k: product.get(k, "") for k in [
            "idProduct", "name", "categoryName", "idCategory", "idExpansion", "idMetacard", "dateAdded"
        ]}
        result["derived"].update({
            "ID Carta ricavato dal nome Cardmarket": card_id,
            "Nome pulito ricavato": clean_name,
            "Colonna prezzo usata dal programma": PRICE_SOURCE_COLUMN,
        })
    else:
        card_id = ""
        clean_name = ""

    if price:
        result["price"] = {k: price.get(k, "") for k in [
            "idProduct", "low", "trend", "avg", "avg1", "avg7", "avg30", "sell", "createdAt"
        ] if k in price}
        # Alcuni JSON hanno createdAt solo a livello root.
        if "createdAt" not in result["price"] and isinstance(price_data, dict):
            result["price"]["priceGuide createdAt"] = price_data.get("createdAt", "")

    df = collection_df.copy()
    if df.empty:
        result["messages"].append("Database collezione vuoto o non caricato.")
    else:
        if "Cardmarket idProduct" not in df.columns:
            df["Cardmarket idProduct"] = ""
        if "ID Carta" not in df.columns:
            df["ID Carta"] = ""

        idp_series = df["Cardmarket idProduct"].apply(normalize_idproduct_for_price_key)
        rows_by_idp = df[idp_series == normalized].copy()
        result["rows_by_idproduct"] = rows_by_idp

        if card_id:
            card_id_series = df["ID Carta"].astype(str).apply(lambda x: normalize_code(x.strip()))
            result["rows_by_card_id"] = df[card_id_series == card_id].copy()

        market_update = {}
        if price:
            def num(key):
                return pd.to_numeric(price.get(key, 0), errors="coerce")
            market_update = {
                "Valore": float(num(PRICE_SOURCE_COLUMN)) if not pd.isna(num(PRICE_SOURCE_COLUMN)) else 0.0,
                "CM_Low": float(num("low")) if not pd.isna(num("low")) else 0.0,
                "CM_Trend": float(num("trend")) if not pd.isna(num("trend")) else 0.0,
                "CM_Avg": float(num("avg")) if not pd.isna(num("avg")) else 0.0,
                "CM_Avg1": float(num("avg1")) if not pd.isna(num("avg1")) else 0.0,
                "CM_Avg7": float(num("avg7")) if not pd.isna(num("avg7")) else 0.0,
                "CM_Avg30": float(num("avg30")) if not pd.isna(num("avg30")) else 0.0,
                "CM_Data_Prezzo": price_data.get("createdAt", "") if isinstance(price_data, dict) else "",
                "Fonte prezzo": f"Cardmarket idProduct {normalized}",
            }
        if product:
            # Diagnostica solamente. Aggiorna valori NON deve riscrivere questo campo se già presente.
            market_update["Cardmarket idProduct"] = normalized
            market_update["Cardmarket Nome"] = clean_name or product.get("name", "")

        preview_rows = []
        target_rows = result["rows_by_idproduct"]
        if not target_rows.empty and market_update:
            for idx, row in target_rows.iterrows():
                for col, new_value in market_update.items():
                    old_value = row.get(col, "")
                    preview_rows.append({
                        "Indice riga": idx,
                        "Campo": col,
                        "Valore attuale": old_value,
                        "Valore da JSON Cardmarket": new_value,
                        "Viene aggiornato?": "Sì" if col != "Cardmarket idProduct" else "No, resta chiave di partenza",
                    })
        result["updates_preview"] = pd.DataFrame(preview_rows)

    if product and price:
        result["ok"] = True
        result["messages"].append("Prodotto e prezzi trovati. Aggiorna valori partirebbe da questo idProduct e toccherebbe solo i campi mercato.")
    return result

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


def _extract_run_summary(script_name, output, returncode, elapsed):
    """Crea un riepilogo breve senza mostrare tutto il log nella dashboard."""
    import re

    label_map = {
        "one_piece_collection_build.py": "Crea da zero",
        "one_piece_collection_update_values.py": "Aggiorna valori",
        "one_piece_collection_sync.py": "Sync espansioni",
    }
    title = label_map.get(script_name, script_name)
    lines = [f"**{title}** completato in **{elapsed}s**." if returncode == 0 else f"**{title}** terminato con errore **{returncode}** dopo **{elapsed}s**."]

    patterns = [
        (r"JSON Cardmarket validati:\s*(.+)", "JSON Cardmarket"),
        (r"Set trovati/preparati:\s*(\d+)", "Set preparati"),
        (r"Carte Bandai raw già presenti:\s*(\d+)", "Carte raw già presenti"),
        (r"Promo P escluse dal database Bandai:\s*(\d+)", "Promo escluse"),
        (r"Valori aggiornati:\s*(\d+)", "Valori aggiornati"),
        (r"Nuove carte aggiunte:\s*(\d+)", "Nuove carte aggiunte"),
        (r"Doppioni rimossi:\s*(\d+)", "Doppioni rimossi"),
        (r"Excel finale:\s*(.+)", "Excel finale"),
        (r"JSON finale:\s*(.+)", "JSON finale"),
        (r"Log run:\s*(.+)", "Log"),
    ]
    found = []
    for pattern, label in patterns:
        m = re.search(pattern, output or "")
        if m:
            value = m.group(1).strip()
            if len(value) > 140:
                value = value[:137] + "..."
            found.append(f"- {label}: `{value}`")

    if found:
        lines.append("\n".join(found[:8]))
    elif returncode == 0:
        lines.append("Operazione completata. I dettagli restano disponibili nella scheda **File e log**.")
    else:
        tail = "\n".join((output or "").splitlines()[-8:])
        if tail:
            lines.append("Ultime righe errore:\n```text\n" + tail[-1200:] + "\n```")
    return "\n\n".join(lines)


def run_script(script_name):
    script_path = SRC / script_name
    if not script_path.exists():
        st.error(f"Script non trovato: {script_path}")
        return False

    st.session_state["last_run_script"] = script_name
    st.session_state["last_run_output"] = ""
    st.session_state["last_run_returncode"] = None
    st.session_state["last_run_summary"] = ""

    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"

    cmd = [sys.executable, "-u", str(script_path)]
    started_at = time.time()

    label_map = {
        "one_piece_collection_build.py": "Creo la collezione da zero",
        "one_piece_collection_update_values.py": "Aggiorno i valori di mercato",
        "one_piece_collection_sync.py": "Sincronizzo nuove espansioni",
    }
    label = label_map.get(script_name, script_name)

    output = ""
    returncode = -1
    try:
        with st.spinner(f"{label}... Attendi il riepilogo finale."):
            completed = subprocess.run(
                cmd,
                cwd=str(ROOT),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=env,
            )
            output = completed.stdout or ""
            returncode = completed.returncode
    except Exception as exc:
        output = f"ERRORE AVVIO SCRIPT: {exc}"
        returncode = -1

    elapsed = int(time.time() - started_at)
    summary = _extract_run_summary(script_name, output, returncode, elapsed)

    st.session_state["last_run_output"] = output
    st.session_state["last_run_returncode"] = returncode
    st.session_state["last_run_summary"] = summary

    if returncode == 0:
        st.toast("Operazione completata", icon="✅")
        return True

    st.toast("Operazione terminata con errore", icon="⚠️")
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


# -----------------------------------------------------------------------------
# Comandi principali: niente sidebar, barra azioni sempre visibile in alto.
# -----------------------------------------------------------------------------
st.markdown('<div class="op-toolbar-marker"></div>', unsafe_allow_html=True)
cmd_cols = st.columns([1.1, 1.1, 1.1, 0.9, 2.4], vertical_alignment="center")
with cmd_cols[0]:
    if st.button("🧱 Crea da zero", use_container_width=True):
        run_script("one_piece_collection_build.py")
with cmd_cols[1]:
    if st.button("💶 Aggiorna valori", use_container_width=True):
        run_script("one_piece_collection_update_values.py")
with cmd_cols[2]:
    if st.button("🔄 Sync espansioni", use_container_width=True):
        run_script("one_piece_collection_sync.py")
with cmd_cols[3]:
    if st.button("🔁 Ricarica", use_container_width=True):
        st.rerun()
with cmd_cols[4]:
    pass

if st.session_state.get("last_run_summary"):
    st.markdown('<div class="op-run-summary">', unsafe_allow_html=True)
    st.markdown(st.session_state["last_run_summary"])
    st.markdown('</div>', unsafe_allow_html=True)

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
    "Usa le schede qui sotto come indice rapido: **Panoramica**, **Top valore**, **Trend prezzi**, **Gestione carte**, **Domande database**, **Debug mode**, **File e log**."
)

tab_overview, tab_top_value, tab_trends, tab_cards, tab_ai, tab_debug, tab_files = st.tabs([
    "📌 Panoramica",
    "💰 Top valore",
    "📈 Trend prezzi",
    "🃏 Gestione carte",
    "🔎 Domande database",
    "🧪 Debug mode",
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
    st.markdown(
        '<div class="op-card">Qui puoi filtrare, scegliere le colonne, modificare i dati, aggiungere righe, eliminare carte e rimuovere duplicati. I valori economici sono in euro (€).</div>',
        unsafe_allow_html=True,
    )

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

    def _is_effectively_empty_row(row, ignore_cols=None):
        ignore_cols = set(ignore_cols or [])
        for col, val in row.items():
            if col in ignore_cols:
                continue
            if pd.isna(val):
                continue
            if isinstance(val, (int, float)) and float(val) == 0:
                continue
            if str(val).strip() != "":
                return False
        return True

    def _normalise_card_id_fields(df):
        out = df.copy()
        if "ID Carta" in out.columns:
            out["ID Carta"] = out["ID Carta"].astype(str).str.strip().str.upper().str.replace(".", "-", regex=False).str.replace("_", "-", regex=False)
        if "Espansione" in out.columns and "ID Carta" in out.columns:
            missing = out["Espansione"].isna() | out["Espansione"].astype(str).str.strip().eq("")
            out.loc[missing, "Espansione"] = out.loc[missing, "ID Carta"].apply(_guess_expansion_from_card_id)
        if "Numero" in out.columns and "ID Carta" in out.columns:
            missing = out["Numero"].isna() | out["Numero"].astype(str).str.strip().eq("")
            out.loc[missing, "Numero"] = out.loc[missing, "ID Carta"].apply(_guess_number_from_card_id)
        return out

    # -------------------------------------------------------------------------
    # Ricerca principale + ricerca avanzata
    # -------------------------------------------------------------------------
    filtered = collection.copy()
    st.subheader("Ricerca")
    main_filter_cols = st.columns([2.2, 1, 1, 1])
    with main_filter_cols[0]:
        search = st.text_input("Cerca", placeholder="Nome, ID carta, variante...", key="cards_search_main")
    with main_filter_cols[1]:
        if "Lingua" in filtered.columns:
            langs = sorted([x for x in filtered["Lingua"].dropna().astype(str).unique() if x.strip()])
            selected_langs = st.multiselect("Lingua", langs, default=[], key="cards_filter_lang")
        else:
            selected_langs = []
    with main_filter_cols[2]:
        if "Espansione" in filtered.columns:
            exps = sorted([x for x in filtered["Espansione"].dropna().astype(str).unique() if x.strip()])
            selected_exps = st.multiselect("Espansione", exps, default=[], key="cards_filter_exp")
        else:
            selected_exps = []
    with main_filter_cols[3]:
        owned_only = st.checkbox("Solo possedute", value=False, key="cards_filter_owned")

    if search:
        s = search.lower().strip()
        mask = pd.Series(False, index=filtered.index)
        for col in ["ID Carta", "Nome", "Variante", "Espansione", "Numero", "Rarità"]:
            if col in filtered.columns:
                mask = mask | filtered[col].astype(str).str.lower().str.contains(s, na=False)
        filtered = filtered[mask]

    if selected_langs and "Lingua" in filtered.columns:
        filtered = filtered[filtered["Lingua"].isin(selected_langs)]
    if selected_exps and "Espansione" in filtered.columns:
        filtered = filtered[filtered["Espansione"].isin(selected_exps)]
    if owned_only and "Quantità" in filtered.columns:
        filtered = filtered[pd.to_numeric(filtered["Quantità"], errors="coerce").fillna(0) > 0]

    with st.expander("Ricerca avanzata", expanded=False):
        adv_cols = st.columns(4)
        with adv_cols[0]:
            if "Rarità" in filtered.columns:
                rarities = sorted([x for x in collection["Rarità"].dropna().astype(str).unique() if x.strip()])
                selected_rarities = st.multiselect("Rarità", rarities, default=[], key="cards_filter_rarity")
            else:
                selected_rarities = []
            if "Tipo carta" in filtered.columns:
                types = sorted([x for x in collection["Tipo carta"].dropna().astype(str).unique() if x.strip()])
                selected_types = st.multiselect("Tipo carta", types, default=[], key="cards_filter_type")
            else:
                selected_types = []
        with adv_cols[1]:
            if "Color" in filtered.columns:
                colors = sorted([x for x in collection["Color"].dropna().astype(str).unique() if x.strip()])
                selected_colors = st.multiselect("Color", colors, default=[], key="cards_filter_color")
            else:
                selected_colors = []
            if "Variante" in filtered.columns:
                variants = sorted([x for x in collection["Variante"].dropna().astype(str).unique() if x.strip()])
                selected_variants = st.multiselect("Variante", variants, default=[], key="cards_filter_variant")
            else:
                selected_variants = []
        with adv_cols[2]:
            min_value = st.number_input("Valore minimo (€)", min_value=0.0, value=0.0, step=0.5, format="%.2f", key="cards_filter_min_value")
            min_qty = st.number_input("Quantità minima", min_value=0, value=0, step=1, key="cards_filter_min_qty")
        with adv_cols[3]:
            if "Trend prezzo" in filtered.columns:
                trends = sorted([x for x in collection["Trend prezzo"].dropna().astype(str).unique() if x.strip()])
                selected_trends = st.multiselect("Trend prezzo", trends, default=[], key="cards_filter_trend")
            else:
                selected_trends = []
            max_rows = st.number_input("Limite carte visualizzate", min_value=10, max_value=10000, value=1000, step=100, key="cards_filter_limit")

    if selected_rarities and "Rarità" in filtered.columns:
        filtered = filtered[filtered["Rarità"].isin(selected_rarities)]
    if selected_types and "Tipo carta" in filtered.columns:
        filtered = filtered[filtered["Tipo carta"].isin(selected_types)]
    if selected_colors and "Color" in filtered.columns:
        filtered = filtered[filtered["Color"].isin(selected_colors)]
    if selected_variants and "Variante" in filtered.columns:
        filtered = filtered[filtered["Variante"].isin(selected_variants)]
    if selected_trends and "Trend prezzo" in filtered.columns:
        filtered = filtered[filtered["Trend prezzo"].isin(selected_trends)]
    if "Valore" in filtered.columns and min_value > 0:
        filtered = filtered[pd.to_numeric(filtered["Valore"], errors="coerce").fillna(0) >= float(min_value)]
    if "Quantità" in filtered.columns and min_qty > 0:
        filtered = filtered[pd.to_numeric(filtered["Quantità"], errors="coerce").fillna(0) >= int(min_qty)]

    st.caption(f"Carte trovate: {len(filtered)} su {len(collection)}")

    # -------------------------------------------------------------------------
    # Scelta colonne e ordine
    # -------------------------------------------------------------------------
    preferred_cols = [
        "ID Carta", "Espansione", "Numero", "Nome", "Lingua", "Variante", "Rarità", "Tipo carta", "Color", "Quantità",
        "Valore", "Valore totale", "Valore precedente", "Variazione valore", "Variazione %", "Trend prezzo",
        "Rarità JP ufficiale", "Rarità JP candidate", "Fonte rarità JP", PRICE_SOURCE_COLUMN,
    ]
    available_cols = [c for c in preferred_cols if c in collection.columns] + [c for c in collection.columns if c not in preferred_cols]
    default_cols = [c for c in preferred_cols if c in collection.columns]

    # Streamlit non permette di modificare session_state di un widget dopo che è stato creato.
    # I pulsanti preset impostano quindi una selezione pendente, applicata prima del multiselect.
    if "cards_cols_pending_preset" in st.session_state:
        pending_cols = st.session_state.pop("cards_cols_pending_preset")
        st.session_state["cards_visible_cols"] = [c for c in pending_cols if c in available_cols]

    # Se la struttura dati cambia tra una versione e l'altra, Streamlit può avere in sessione
    # colonne che non esistono più nel dataset corrente. Il multiselect non accetta default
    # fuori dalle options, quindi ripuliamo prima di creare il widget.
    current_cols = st.session_state.get("cards_visible_cols", default_cols)
    if isinstance(current_cols, str):
        current_cols = [current_cols]
    current_cols = [c for c in current_cols if c in available_cols]
    if not current_cols:
        current_cols = default_cols
    st.session_state["cards_visible_cols"] = current_cols

    with st.expander("Colonne visibili e ordine", expanded=False):
        preset_cols = st.columns(4)
        if preset_cols[0].button("Preset base", use_container_width=True):
            st.session_state["cards_cols_pending_preset"] = [c for c in ["ID Carta", "Nome", "Lingua", "Variante", "Rarità", "Quantità", "Valore", "Valore totale"] if c in collection.columns]
            st.rerun()
        if preset_cols[1].button("Preset prezzi", use_container_width=True):
            st.session_state["cards_cols_pending_preset"] = [c for c in ["ID Carta", "Nome", "Quantità", "Valore", "Valore totale", "Valore precedente", "Variazione valore", "Variazione %", "Trend prezzo"] if c in collection.columns]
            st.rerun()
        if preset_cols[2].button("Preset completo", use_container_width=True):
            st.session_state["cards_cols_pending_preset"] = available_cols
            st.rerun()
        if preset_cols[3].button("Reset ordine", use_container_width=True):
            st.session_state["cards_cols_pending_preset"] = default_cols
            st.rerun()

        selected_cols = st.multiselect(
            "Scegli colonne e ordine di visualizzazione",
            options=available_cols,
            default=st.session_state.get("cards_visible_cols", default_cols),
            key="cards_visible_cols",
            help="L'ordine selezionato qui sarà l'ordine della tabella modificabile.",
        )

    show_cols = selected_cols or default_cols

    # -------------------------------------------------------------------------
    # Azioni rapide: aggiunta riga e rimozione duplicati
    # -------------------------------------------------------------------------
    st.subheader("Tabella carte")

    action_cols = st.columns([1.1, 1.2, 1.4, 2.8])
    if "new_rows_to_show" not in st.session_state:
        st.session_state["new_rows_to_show"] = 0

    with action_cols[0]:
        if st.button("➕ Aggiungi riga", use_container_width=True):
            st.session_state["new_rows_to_show"] += 1
            st.rerun()
    with action_cols[1]:
        if st.button("🧹 Elimina doppioni identici", use_container_width=True):
            before = len(collection)
            deduped = collection.drop_duplicates(keep="first").reset_index(drop=True)
            removed = before - len(deduped)
            if removed > 0:
                with st.status("Rimozione doppioni e salvataggio in corso...", expanded=True) as status:
                    st.write(f"Doppioni identici trovati: {removed}")
                    save_collection_from_streamlit(deduped, source="streamlit_remove_exact_duplicates")
                    status.update(label=f"Doppioni rimossi: {removed}", state="complete")
                st.cache_data.clear()
                time.sleep(0.5)
                st.rerun()
            else:
                st.info("Nessun doppione identico trovato.")
    with action_cols[2]:
        if st.button("Annulla nuove righe non salvate", use_container_width=True, disabled=st.session_state.get("new_rows_to_show", 0) == 0):
            st.session_state["new_rows_to_show"] = 0
            st.rerun()
    with action_cols[3]:
        st.empty()

    # -------------------------------------------------------------------------
    # Data editor con eliminazione tramite selezione riga e nuove righe compilabili
    # -------------------------------------------------------------------------
    filtered_limited = filtered.head(int(max_rows)).copy() if "max_rows" in locals() else filtered.copy()
    editor_df = filtered_limited[show_cols].copy() if show_cols else filtered_limited.copy()

    # Forza tipi editabili nella tabella: Streamlit può bloccare colonne con tipi misti.
    numeric_cols = {"Quantità", "Valore", "Valore totale", "Valore precedente", "Variazione valore", "Variazione %", "CM_Low", "CM_Trend", "CM_Avg", "CM_Avg1", "CM_Avg7", "CM_Avg30", "Power", "Counter", "Cost", "Life", "Cardmarket Prodotti per carta"}
    for _col in list(editor_df.columns):
        if _col in numeric_cols:
            editor_df[_col] = pd.to_numeric(editor_df[_col], errors="coerce").fillna(0)
        else:
            editor_df[_col] = editor_df[_col].fillna("").astype(str)

    editor_df.insert(0, "_new_row", False)
    editor_df.insert(0, "_row_id", filtered_limited.index.astype(int))

    # Aggiunge righe vuote compilabili direttamente nella tabella.
    new_rows = int(st.session_state.get("new_rows_to_show", 0) or 0)
    if new_rows > 0:
        blank_rows = []
        base_cols = list(editor_df.columns)
        for i in range(new_rows):
            row = {c: "" for c in base_cols}
            row["_row_id"] = -1 - i
            row["_new_row"] = True
            if "Lingua" in row:
                row["Lingua"] = "EN"
            if "Variante" in row:
                row["Variante"] = "Base"
            if "Quantità" in row:
                row["Quantità"] = 0
            if "Valore" in row:
                row["Valore"] = 0.0
            if "Valore totale" in row:
                row["Valore totale"] = 0.0
            if PRICE_SOURCE_COLUMN in row:
                row[PRICE_SOURCE_COLUMN] = "Manuale"
            blank_rows.append(row)
        editor_df = pd.concat([editor_df, pd.DataFrame(blank_rows)], ignore_index=True)

    column_config = {
        "_row_id": None,
        "_new_row": None,
    }
    for _col in show_cols:
        if _col == "Quantità":
            column_config[_col] = st.column_config.NumberColumn("Quantità", min_value=0, step=1, format="%d", help="Quante copie possiedi.")
        elif _col in {"Valore", "Valore totale", "Valore precedente", "Variazione valore", "CM_Low", "CM_Trend", "CM_Avg", "CM_Avg1", "CM_Avg7", "CM_Avg30"}:
            label = f"{_col} (€)" if not _col.startswith("CM_") else _col
            column_config[_col] = money_column(label) if not _col.startswith("CM_") else st.column_config.NumberColumn(_col, format="%.2f")
        elif _col == "Variazione %":
            column_config[_col] = st.column_config.NumberColumn("Variazione %", format="%.2f%%")
        elif _col in {"Power", "Counter", "Cost", "Life", "Cardmarket Prodotti per carta"}:
            column_config[_col] = st.column_config.NumberColumn(_col, step=1, format="%d")
        else:
            column_config[_col] = st.column_config.TextColumn(_col)

    edited_df = st.data_editor(
        editor_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        disabled=["_row_id", "_new_row"],
        column_config=column_config,
        key="cards_full_editor_v2",
        height=620,
    )

    # Elimina carte tramite selezione riga su una tabella compatta separata.
    selected_delete_ids = []
    with st.expander("Elimina carte", expanded=False):
        delete_cols = [c for c in ["ID Carta", "Nome", "Lingua", "Rarità", "Variante", "Espansione", "Numero", "Quantità", "Valore"] if c in filtered_limited.columns]
        delete_map = filtered_limited[delete_cols].copy() if delete_cols else filtered_limited.copy()
        delete_map.insert(0, "_row_id", filtered_limited.index.astype(int))
        try:
            delete_event = st.dataframe(
                delete_map.drop(columns=["_row_id"]),
                use_container_width=True,
                hide_index=True,
                height=280,
                selection_mode="multi-row",
                on_select="rerun",
                key="cards_delete_selection_table",
            )
            selected_positions = list(getattr(delete_event, "selection", {}).get("rows", [])) if delete_event is not None else []
        except TypeError:
            labels = []
            label_to_id = {}
            for pos, (_, r) in enumerate(delete_map.iterrows()):
                label = " | ".join(str(r.get(c, "")) for c in delete_cols[:5])
                label = f"{pos + 1}. {label}"
                labels.append(label)
                label_to_id[label] = int(r["_row_id"])
            chosen = st.multiselect("Seleziona carte da eliminare", labels, key="cards_delete_selection_fallback")
            selected_positions = []
            selected_delete_ids = [label_to_id[x] for x in chosen]

        if not selected_delete_ids and selected_positions:
            for pos in selected_positions:
                if 0 <= int(pos) < len(delete_map):
                    selected_delete_ids.append(int(delete_map.iloc[int(pos)]["_row_id"]))

        if selected_delete_ids:
            st.warning(f"Carte selezionate per eliminazione: {len(selected_delete_ids)}")

    # Calcola modifiche non salvate, nuove righe e selezioni da eliminare.
    changes = 0
    new_rows_to_save = []

    for _, row in edited_df.iterrows():
        row_id = int(row.get("_row_id", -999999))
        is_new = bool(row.get("_new_row", False)) or row_id < 0

        row_payload = {c: row.get(c, "") for c in show_cols if c in edited_df.columns}
        if is_new:
            if not _is_effectively_empty_row(row_payload):
                new_rows_to_save.append(row_payload)
            continue

        if row_id not in collection.index:
            continue
        for col in show_cols:
            if col not in collection.columns or col not in edited_df.columns:
                continue
            new_value = row.get(col, "")
            old_value = collection.at[row_id, col]
            if col in numeric_cols:
                new_num = pd.to_numeric(new_value, errors="coerce")
                old_num = pd.to_numeric(old_value, errors="coerce")
                new_cmp = 0 if pd.isna(new_num) else (int(new_num) if col == "Quantità" else float(new_num))
                old_cmp = 0 if pd.isna(old_num) else (int(old_num) if col == "Quantità" else float(old_num))
                if new_cmp != old_cmp:
                    changes += 1
            else:
                new_cmp = "" if pd.isna(new_value) else str(new_value)
                old_cmp = "" if pd.isna(old_value) else str(old_value)
                if new_cmp != old_cmp:
                    changes += 1

    has_pending = changes > 0 or bool(new_rows_to_save) or bool(selected_delete_ids)
    if has_pending:
        st.markdown(
            f'<div class="op-warning">⚠️ Modifiche non salvate: {changes} campi modificati, {len(new_rows_to_save)} nuove carte compilate, {len(selected_delete_ids)} carte selezionate per eliminazione.</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown('<div class="op-ok">Nessuna modifica da salvare.</div>', unsafe_allow_html=True)

    save_col, delete_col, hint_col = st.columns([1.2, 1.4, 3])
    with save_col:
        save_pressed = st.button("💾 Salva modifiche", type="primary", use_container_width=True, disabled=not has_pending)
    with delete_col:
        delete_pressed = st.button("🗑️ Elimina carte selezionate", use_container_width=True, disabled=len(selected_delete_ids) == 0)
    with hint_col:
        st.caption("Il salvataggio aggiorna JSON ed Excel, con backup automatico.")

    def _apply_editor_changes(save_deletions=True, save_field_changes=True, save_new_rows=True):
        updated = collection.copy()
        local_changes = 0

        if save_deletions and selected_delete_ids:
            updated = updated.drop(index=selected_delete_ids, errors="ignore")
            local_changes += len(selected_delete_ids)

        if save_field_changes:
            for _, row in edited_df.iterrows():
                row_id = int(row.get("_row_id", -999999))
                is_new = bool(row.get("_new_row", False)) or row_id < 0
                if is_new or row_id not in updated.index:
                    continue
                for col in show_cols:
                    if col not in updated.columns or col not in edited_df.columns:
                        continue
                    new_value = row.get(col, "")
                    old_value = updated.at[row_id, col]
                    if col in numeric_cols:
                        new_num = pd.to_numeric(new_value, errors="coerce")
                        new_value = 0 if pd.isna(new_num) else (int(new_num) if col == "Quantità" else float(new_num))
                        old_num = pd.to_numeric(old_value, errors="coerce")
                        old_cmp = 0 if pd.isna(old_num) else (int(old_num) if col == "Quantità" else float(old_num))
                        changed = new_value != old_cmp
                    else:
                        new_value = "" if pd.isna(new_value) else str(new_value)
                        old_cmp = "" if pd.isna(old_value) else str(old_value)
                        changed = new_value != old_cmp
                    if changed:
                        updated.at[row_id, col] = new_value
                        local_changes += 1

        if save_new_rows and new_rows_to_save:
            for payload in new_rows_to_save:
                for col in payload.keys():
                    if col not in updated.columns:
                        updated[col] = 0 if col in numeric_cols else ""
                row = {col: "" for col in updated.columns}
                for col, value in payload.items():
                    if col in row:
                        row[col] = value
                if "ID Carta" in row:
                    row["ID Carta"] = str(row.get("ID Carta", "")).strip().upper().replace(".", "-").replace("_", "-")
                if "Espansione" in row and not str(row.get("Espansione", "")).strip():
                    row["Espansione"] = _guess_expansion_from_card_id(row.get("ID Carta", ""))
                if "Numero" in row and not str(row.get("Numero", "")).strip():
                    row["Numero"] = _guess_number_from_card_id(row.get("ID Carta", ""))
                if PRICE_SOURCE_COLUMN in row and not str(row.get(PRICE_SOURCE_COLUMN, "")).strip():
                    row[PRICE_SOURCE_COLUMN] = "Manuale"
                updated = pd.concat([updated, pd.DataFrame([row])], ignore_index=True)
                local_changes += 1

        updated = _normalise_card_id_fields(updated).reset_index(drop=True)
        if "Quantità" in updated.columns:
            updated["Quantità"] = pd.to_numeric(updated["Quantità"], errors="coerce").fillna(0).clip(lower=0).round(0).astype(int)
        if "Valore" in updated.columns:
            updated["Valore"] = pd.to_numeric(updated["Valore"], errors="coerce").fillna(0)
        if "Valore totale" not in updated.columns and "Quantità" in updated.columns and "Valore" in updated.columns:
            updated["Valore totale"] = updated["Quantità"] * updated["Valore"]
        return updated, local_changes

    if delete_pressed:
        try:
            with st.status("Eliminazione e salvataggio in corso...", expanded=True) as status:
                st.write(f"Carte selezionate da eliminare: {len(selected_delete_ids)}")
                updated, local_changes = _apply_editor_changes(save_deletions=True, save_field_changes=False, save_new_rows=False)
                save_collection_from_streamlit(updated, source="streamlit_delete_selected_rows")
                status.update(label="Eliminazione completata", state="complete")
            st.success(f"Carte eliminate: {len(selected_delete_ids)}")
            st.cache_data.clear()
            time.sleep(0.5)
            st.rerun()
        except Exception as exc:
            st.error(f"Errore durante l'eliminazione: {exc}")

    if save_pressed:
        try:
            with st.status("Salvataggio modifiche in corso...", expanded=True) as status:
                st.write("Creo backup automatico dei file finali.")
                st.write("Aggiorno JSON finale.")
                st.write("Rigenero Excel con Dashboard.")
                updated, local_changes = _apply_editor_changes(save_deletions=True, save_field_changes=True, save_new_rows=True)
                save_collection_from_streamlit(updated, source="streamlit_table_edit")
                status.update(label="Salvataggio completato", state="complete")
            st.session_state["new_rows_to_show"] = 0
            st.success(f"Modifiche salvate. Operazioni applicate: {local_changes}. Excel e JSON aggiornati.")
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


with tab_debug:
    st.header("Debug mode")
    st.markdown("Verifica cosa farebbe **Aggiorna valori** partendo da un singolo `Cardmarket idProduct`, senza modificare file.")

    dbg_cols = st.columns([1.2, 0.8, 2.0])
    with dbg_cols[0]:
        debug_idp = st.text_input("Cardmarket idProduct", value="", placeholder="es. 1234567")
    with dbg_cols[1]:
        run_debug = st.button("Analizza idProduct", use_container_width=True)
    with dbg_cols[2]:
        st.caption("La simulazione legge i JSON in `json/`, cerca prodotto e prezzo, trova le righe nel database e mostra i campi mercato che verrebbero aggiornati.")

    if run_debug:
        debug_result = debug_cardmarket_update_for_idproduct(debug_idp, collection)

        if debug_result["ok"]:
            st.success(debug_result["messages"][-1])
        else:
            for msg in debug_result["messages"]:
                st.warning(msg)

        step1, step2, step3, step4, step5 = st.tabs([
            "1. File JSON", "2. Prodotto", "3. Prezzi", "4. Match database", "5. Aggiornamento simulato"
        ])

        with step1:
            st.subheader("File usati")
            if debug_result["files"]:
                st.dataframe(pd.DataFrame([{"Tipo": k, "Percorso": v} for k, v in debug_result["files"].items()]), use_container_width=True, hide_index=True)
            else:
                st.info("Nessun file JSON individuato.")
            st.subheader("Dati ricavati")
            st.dataframe(pd.DataFrame([{"Campo": k, "Valore": v} for k, v in debug_result["derived"].items()]), use_container_width=True, hide_index=True)

        with step2:
            st.subheader("Riga products_singles")
            if debug_result["product"]:
                st.dataframe(pd.DataFrame([debug_result["product"]]), use_container_width=True, hide_index=True)
            else:
                st.error("Nessun prodotto trovato per questo idProduct.")

        with step3:
            st.subheader("Riga price_guide")
            if debug_result["price"]:
                st.dataframe(pd.DataFrame([debug_result["price"]]), use_container_width=True, hide_index=True)
            else:
                st.error("Nessun prezzo trovato per questo idProduct.")

        with step4:
            st.subheader("Righe trovate usando Cardmarket idProduct")
            rows_idp = debug_result["rows_by_idproduct"]
            if isinstance(rows_idp, pd.DataFrame) and not rows_idp.empty:
                show_cols = [c for c in ["ID Carta", "Nome", "Lingua", "Variante", "Rarità", "Quantità", "Valore", "Cardmarket idProduct", "Cardmarket Nome"] if c in rows_idp.columns]
                st.dataframe(rows_idp[show_cols], use_container_width=True)
            else:
                st.warning("Nessuna riga del database usa questo Cardmarket idProduct.")

            st.subheader("Righe trovate usando l'ID Carta ricavato")
            rows_card = debug_result["rows_by_card_id"]
            if isinstance(rows_card, pd.DataFrame) and not rows_card.empty:
                show_cols = [c for c in ["ID Carta", "Nome", "Lingua", "Variante", "Rarità", "Quantità", "Valore", "Cardmarket idProduct", "Cardmarket Nome"] if c in rows_card.columns]
                st.dataframe(rows_card[show_cols], use_container_width=True)
            else:
                st.info("Nessuna riga trovata per ID Carta ricavato, oppure ID Carta non ricavabile dal nome prodotto.")

        with step5:
            st.subheader("Campi che verrebbero aggiornati")
            preview = debug_result["updates_preview"]
            if isinstance(preview, pd.DataFrame) and not preview.empty:
                st.dataframe(preview, use_container_width=True, hide_index=True)
                st.info("Nota: `Cardmarket idProduct` viene mostrato solo per diagnosi, ma non viene sovrascritto durante Aggiorna valori se è già la chiave di partenza.")
            else:
                st.warning("Nessun aggiornamento simulabile: manca il match nel database o manca la riga prezzo.")

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
