import os
import re
import pandas as pd
from openpyxl import load_workbook

from one_piece_common import *

KEEP_MANUAL_JP_IF_NO_PRICE = True


def variant_rank_from_text(text):
    if not isinstance(text, str):
        return None
    t = text.lower()
    if "base" in t:
        return 1
    m = re.search(r"(?:parallel|alt|reprint).*?(\d+)", t)
    if m:
        return int(m.group(1)) + 1
    m = re.search(r"variante\s+(\d+)", t)
    if m:
        return int(m.group(1))
    return None


def get_header_map(ws):
    return {str(ws.cell(row=1, column=c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(row=1, column=c).value}


def ensure_column(ws, headers, name):
    if name in headers:
        return headers[name]
    col = ws.max_column + 1
    ws.cell(row=1, column=col).value = name
    headers[name] = col
    return col


def get_cell(ws, row, headers, name):
    col = headers.get(name)
    if not col:
        return ""
    v = ws.cell(row=row, column=col).value
    return "" if v is None else v


def set_cell(ws, row, headers, name, value):
    col = ensure_column(ws, headers, name)
    ws.cell(row=row, column=col).value = value


def find_match(row_data, df_cm, by_idproduct):
    card_id = normalize_code(str(row_data.get("ID Carta", "")).strip())
    lang = normalize_language(str(row_data.get("Lingua", "EN")).strip())
    existing_idproduct = str(row_data.get("Cardmarket idProduct", "")).strip()
    variant = str(row_data.get("Variante", "")).strip()
    if not card_id or card_id.startswith("DON-"):
        return None
    if existing_idproduct and existing_idproduct.lower() != "nan":
        try:
            idp = int(float(existing_idproduct))
            if idp in by_idproduct:
                return by_idproduct[idp]
        except Exception:
            pass
    candidates = df_cm[(df_cm["ID Carta"].astype(str) == card_id) & (df_cm["CM Expansion Language"].astype(str) == lang)].copy()
    if candidates.empty:
        return None
    rank = variant_rank_from_text(variant)
    if rank is not None:
        exact = candidates[candidates["Cardmarket Variante N"] == rank]
        if not exact.empty:
            return exact.iloc[0].to_dict()
    return candidates.iloc[0].to_dict()


def apply_number_formats(ws, headers):
    for name in ["Valore", "Valore totale", "CM_Low", "CM_Trend", "CM_Avg", "CM_Avg1", "CM_Avg7", "CM_Avg30"]:
        col = headers.get(name)
        if col:
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=col).number_format = '€ #,##0.00'
    numero_col = headers.get("Numero")
    if numero_col:
        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=numero_col).number_format = "@"
            ws.cell(row=row, column=numero_col).value = format_number_text(ws.cell(row=row, column=numero_col).value)
    qty_col = headers.get("Quantità")
    if qty_col:
        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=qty_col).number_format = "0"


def workbook_to_dataframe(wb):
    ws = wb["Carte"]
    headers = [str(ws.cell(row=1, column=c).value).strip() if ws.cell(row=1, column=c).value else f"Column {c}" for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(2, ws.max_row + 1):
        item = {}
        empty = True
        for c, h in enumerate(headers, start=1):
            v = ws.cell(row=r, column=c).value
            if v not in [None, ""]:
                empty = False
            item[h] = "" if v is None else v
        if not empty:
            rows.append(item)
    return pd.DataFrame(rows)


def update_excel_values():
    start_run_logging("update_values")
    ensure_dirs()
    warn_out_extra_files()
    if not os.path.exists(OUTPUT_XLSX):
        raise FileNotFoundError(f"Excel non trovato: {OUTPUT_XLSX}")
    df_cm = load_cardmarket_prices(CARDMARKET_UPDATED_CSV)
    by_idproduct = {}
    for _, r in df_cm.iterrows():
        try:
            by_idproduct[int(r["idProduct"])] = r.to_dict()
        except Exception:
            pass
    backup_file(OUTPUT_XLSX, move=False)
    backup_file(OUTPUT_JSON, move=False)
    wb = load_workbook(OUTPUT_XLSX)
    if "Carte" not in wb.sheetnames:
        raise ValueError("Nel file Excel non trovo il foglio 'Carte'.")
    ws = wb["Carte"]
    headers = get_header_map(ws)
    required = ["ID Carta", "Lingua", "Variante", "Valore", "Fonte prezzo", "CM_Data_Prezzo", "Cardmarket idProduct", "Cardmarket Nome", "Cardmarket Prodotti per carta", "CM_Low", "CM_Trend", "CM_Avg", "CM_Avg1", "CM_Avg7", "CM_Avg30", "CM Expansion Name", "CM Expansion Code", "CM Expansion Language", "CM Expansion Product", "CM Product Type"]
    for c in required:
        ensure_column(ws, headers, c)
    price_col = PRICE_SOURCE_COLUMN if PRICE_SOURCE_COLUMN in df_cm.columns else "trend"
    report = []
    updated = not_found = skipped_don = skipped_manual_jp = 0
    for row_idx in range(2, ws.max_row + 1):
        row_data = {name: get_cell(ws, row_idx, headers, name) for name in headers}
        card_id = normalize_code(str(row_data.get("ID Carta", "")).strip())
        lang = normalize_language(str(row_data.get("Lingua", "EN")).strip())
        old_value = get_cell(ws, row_idx, headers, "Valore")
        old_idp = get_cell(ws, row_idx, headers, "Cardmarket idProduct")
        if not card_id:
            continue
        if card_id.startswith("DON-"):
            skipped_don += 1
            report.append({"Excel row": row_idx, "ID Carta": card_id, "Lingua": lang, "Status": "SKIPPED_DON", "Old Value": old_value, "New Value": old_value, "Old idProduct": old_idp, "New idProduct": ""})
            continue
        match = find_match(row_data, df_cm, by_idproduct)
        if match is None:
            if lang == "JP" and KEEP_MANUAL_JP_IF_NO_PRICE:
                skipped_manual_jp += 1
                status = "SKIPPED_MANUAL_JP_NO_PRICE"
            else:
                not_found += 1
                status = "NOT_FOUND"
            report.append({"Excel row": row_idx, "ID Carta": card_id, "Lingua": lang, "Status": status, "Old Value": old_value, "New Value": old_value, "Old idProduct": old_idp, "New idProduct": ""})
            continue
        new_value = match.get(price_col, 0)
        if pd.isna(new_value):
            new_value = 0
        idp = match.get("idProduct", "")
        try:
            idp_text = int(idp)
        except Exception:
            idp_text = idp
        set_cell(ws, row_idx, headers, "Valore", float(new_value))
        set_cell(ws, row_idx, headers, "Fonte prezzo", f"Cardmarket {price_col} per idProduct {idp_text} [{lang}]")
        set_cell(ws, row_idx, headers, "CM_Data_Prezzo", match.get("Cardmarket Price Created At", ""))
        set_cell(ws, row_idx, headers, "Cardmarket idProduct", idp_text)
        set_cell(ws, row_idx, headers, "Cardmarket Nome", match.get("Cardmarket Nome", ""))
        set_cell(ws, row_idx, headers, "Cardmarket Prodotti per carta", int(match.get("Cardmarket Prodotti per carta", 0)))
        set_cell(ws, row_idx, headers, "CM_Low", clean_numeric(match.get("low", "")))
        set_cell(ws, row_idx, headers, "CM_Trend", clean_numeric(match.get("trend", "")))
        set_cell(ws, row_idx, headers, "CM_Avg", clean_numeric(match.get("avg", "")))
        set_cell(ws, row_idx, headers, "CM_Avg1", clean_numeric(match.get("avg1", "")))
        set_cell(ws, row_idx, headers, "CM_Avg7", clean_numeric(match.get("avg7", "")))
        set_cell(ws, row_idx, headers, "CM_Avg30", clean_numeric(match.get("avg30", "")))
        set_cell(ws, row_idx, headers, "CM Expansion Name", match.get("CM Expansion Name", ""))
        set_cell(ws, row_idx, headers, "CM Expansion Code", match.get("CM Expansion Code", ""))
        set_cell(ws, row_idx, headers, "CM Expansion Language", match.get("CM Expansion Language", lang))
        set_cell(ws, row_idx, headers, "CM Expansion Product", match.get("CM Expansion Product", ""))
        set_cell(ws, row_idx, headers, "CM Product Type", match.get("CM Product Type", "Standard"))
        current_variant = str(row_data.get("Variante", "")).strip().lower()
        if current_variant == "" or "non trovata" in current_variant or "manuale" in current_variant:
            set_cell(ws, row_idx, headers, "Variante", match.get("Variante", ""))
        updated += 1
        report.append({"Excel row": row_idx, "ID Carta": card_id, "Lingua": lang, "Status": "UPDATED", "Old Value": old_value, "New Value": float(new_value), "Old idProduct": old_idp, "New idProduct": idp_text})
    apply_number_formats(ws, headers)
    # Converto il foglio aggiornato in DataFrame, calcolo trend contro il JSON precedente
    # e ricreo l'Excel completo con Dashboard/KPI/grafici.
    final_df = workbook_to_dataframe(wb)
    final_df = add_price_trends_from_latest_backup(final_df)
    save_price_trend_reports(final_df)
    final_df.to_csv(UPDATED_STG_CSV, index=False, encoding="utf-8-sig")
    save_final_json_from_df(final_df, "one_piece_collection_update_values.py")
    create_collection_workbook_with_dashboard(final_df, OUTPUT_XLSX)
    pd.DataFrame(report).to_csv(UPDATE_REPORT_CSV, index=False, encoding="utf-8-sig")
    print("\nAggiornamento completato.")
    print(f"Righe aggiornate: {updated}")
    print(f"Non trovate: {not_found}")
    print(f"JP manuali senza prezzo lasciate intatte: {skipped_manual_jp}")
    print(f"DON saltate: {skipped_don}")
    print("File finali in out/:")
    print(f"- {OUTPUT_XLSX}")
    print(f"- {OUTPUT_JSON}")
    print("Intermedi in stg/.")

if __name__ == "__main__":
    update_excel_values()
