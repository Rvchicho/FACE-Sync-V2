import re
import os
import datetime
import pandas as pd

def limpiar_monto_boliviano(val) -> float:
    if pd.isna(val) or val is None:
        return 0.0
    val_str = str(val).strip()
    if not val_str:
        return 0.0
    if '.' in val_str and ',' in val_str:
        if val_str.rfind(',') > val_str.rfind('.'):
            val_str = val_str.replace('.', '').replace(',', '.')
        else:
            val_str = val_str.replace(',', '')
    elif ',' in val_str:
        val_str = val_str.replace(',', '.')
    try:
        return float(val_str)
    except ValueError:
        return 0.0


def normalizar_codigo(val) -> str:
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if '.' in val_str:
        val_str = val_str.split('.')[0]
    return val_str


def parsear_glosa_inteligente(glosa_raw: str) -> dict:
    if pd.isna(glosa_raw) or not str(glosa_raw).strip():
        return {
            "payee": "SIN PROVEEDOR",
            "invoice_no": "F-0",
            "description": "SIN DESCRIPCION",
            "payment_ref": "N/A",
            "activity_id": "0"
        }

    glosa = str(glosa_raw).strip()
    partes = [p.strip() for p in glosa.split('/') if p.strip()]

    payee, invoice_no, description, payment_ref, activity_id = None, None, None, None, None
    partes_remanentes = []

    for p in partes:
        p_upper = p.upper()
        if re.match(r'^(ACT|ACTIVIDAD|ACT_ID|A)[\:\s\-]', p_upper):
            val = re.sub(r'^(ACT|ACTIVIDAD|ACT_ID|A)[\:\s\-]', '', p, flags=re.IGNORECASE).strip()
            activity_id = normalizar_codigo(val)
        elif re.match(r'^(F|FACT|FACTURA)[\:\s\-]', p_upper):
            val = re.sub(r'^(F|FACT|FACTURA)[\:\s\-]', '', p, flags=re.IGNORECASE).strip()
            invoice_no = f"F-{val}" if not val.upper().startswith("F-") else val.upper()
        elif re.match(r'^(CH|CHEQUE|TR|TRANSF|REF)[\:\s\-]', p_upper):
            val = re.sub(r'^(CH|CHEQUE|TR|TRANSF|REF)[\:\s\-]', '', p, flags=re.IGNORECASE).strip()
            payment_ref = val.upper() if val.upper().startswith("CH") else f"CH-{val}"
        elif re.match(r'^(D|DESC|GLOSA|CONCEPTO)[\:\s\-]', p_upper):
            description = re.sub(r'^(D|DESC|GLOSA|CONCEPTO)[\:\s\-]', '', p, flags=re.IGNORECASE).strip()
        elif re.match(r'^(P|PROV|PROVEEDOR|PAYEE)[\:\s\-]', p_upper):
            payee = re.sub(r'^(P|PROV|PROVEEDOR|PAYEE)[\:\s\-]', '', p, flags=re.IGNORECASE).strip().upper()
        else:
            partes_remanentes.append(p)

    partes_no_clasificadas = []
    for p in partes_remanentes:
        p_upper = p.upper()
        if not activity_id and p.isdigit() and 1 <= int(p) <= 200:
            activity_id = normalizar_codigo(p)
        elif not invoice_no and re.search(r'(?:FACT|FAC|F)[\.\-\s]*(\d+)', p_upper):
            match = re.search(r'(?:FACT|FAC|F)[\.\-\s]*(\d+)', p_upper)
            invoice_no = f"F-{match.group(1)}"
        elif not payment_ref and re.search(r'^(?:CH|CHEQUE|TR|TRANSF|REF)[\.\-\s]*(\d+)', p_upper):
            match = re.search(r'^(?:CH|CHEQUE|TR|TRANSF|REF)[\.\-\s]*(\d+)', p_upper)
            payment_ref = f"CH-{match.group(1)}"
        else:
            partes_no_clasificadas.append(p)

    if partes_no_clasificadas:
        if not payee:
            payee = partes_no_clasificadas.pop(0).upper()
        if not description and partes_no_clasificadas:
            description = " / ".join(partes_no_clasificadas)

    return {
        "payee": payee if payee else "SIN PROVEEDOR",
        "invoice_no": invoice_no if invoice_no else "F-0",
        "description": description if description else (payee if payee else "SIN DESCRIPCION"),
        "payment_ref": payment_ref if payment_ref else "N/A",
        "activity_id": activity_id if activity_id else "0"
    }


def cargar_matriz_budget(file_budget_input=None) -> tuple[dict, pd.DataFrame]:
    """Busca el archivo de presupuesto local en múltiples rutas posibles."""
    mapa_cruce = {}
    df_b = pd.DataFrame()
    
  rutas_posibles = [
        'FACE MALARIA_pestaña Budget.csv',
        os.path.join('data', 'FACE MALARIA_pestaña Budget.csv'),
        'Budget.csv',
        os.path.join('data', 'Budget.csv')
    ]
    
    try:
        if file_budget_input is not None:
            df_b = pd.read_csv(file_budget_input, sep=';', skiprows=5, encoding='utf-8')
        else:
            for ruta in rutas_posibles:
                if os.path.exists(ruta):
                    df_b = pd.read_csv(ruta, sep=';', skiprows=5, encoding='utf-8')
                    break
            
        if not df_b.empty:
            df_b = df_b.dropna(how='all', axis=1).dropna(how='all', axis=0)
            
            for _, row in df_b.iterrows():
                act_id = normalizar_codigo(row.get('Activity ID', ''))
                acc = normalizar_codigo(row.get('UNDP Budget account', ''))
                b_line = normalizar_codigo(row.get('Budget Line #', ''))
                
                if act_id and acc and b_line:
                    mapa_cruce[(act_id, acc)] = b_line
    except Exception:
        pass
        
    return mapa_cruce, df_b


def cargar_mayor_presupuestario(file_bytes, file_budget_input=None) -> tuple[pd.DataFrame, pd.DataFrame]:
    mapa_cruce, df_budget_cat = cargar_matriz_budget(file_budget_input)

    df_raw = pd.read_excel(file_bytes, header=17)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    
    cols = list(df_raw.columns)
    col_codigo = cols[0] if len(cols) > 0 else "Ent."
    col_cbte = cols[2] if len(cols) > 2 else "Cbte"
    col_fecha = cols[3] if len(cols) > 3 else "Fecha"
    col_glosa = cols[4] if len(cols) > 4 else "Glosa"
    col_importe = cols[7] if len(cols) > 7 else "Importe"

    registros = []
    current_account = "71610"

    for idx, row in df_raw.iterrows():
        texto_col_a = str(row.get(col_codigo, '')).strip()
        if texto_col_a:
            match_acc = re.search(r'\b(\d{5})\b', texto_col_a)
            if match_acc:
                current_account = match_acc.group(1)

        voucher = str(row.get(col_cbte, '')).strip()
        fecha_raw = row.get(col_fecha, '')
        glosa = row.get(col_glosa, '')
        monto_raw = row.get(col_importe, 0)
        
        monto = limpiar_monto_boliviano(monto_raw)
        if pd.isna(glosa) or str(glosa).startswith("---") or monto <= 0:
            continue
            
        parsed = parsear_glosa_inteligente(glosa)
        act_id_val = normalizar_codigo(parsed["activity_id"])
        acc_val = normalizar_codigo(current_account)
        
        budget_line_calculado = mapa_cruce.get((act_id_val, acc_val), "0")

        fecha_dt = pd.to_datetime(fecha_raw, errors='coerce')
        fecha_val = fecha_dt.date() if pd.notna(fecha_dt) else str(fecha_raw).strip()

        expenses_87 = round(monto * 0.87, 2)
        vat_13 = round(monto * 0.13, 2)
        
        registros.append({
            "Voucher No": voucher,
            "Fecha": fecha_val,
            "Account": acc_val,
            "Activity ID": act_id_val,
            "Payee / Vendor": parsed["payee"],
            "Invoice No": parsed["invoice_no"],
            "Description": parsed["description"],
            "Payment / Ref": parsed["payment_ref"],
            "Budget Line No": budget_line_calculado,
            "Payment (Bs)": monto,
            "Expenses (87%)": expenses_87,
            "VAT Tax (13%)": vat_13,
            "Glosa Original": glosa
        })
        
    df_res = pd.DataFrame(registros)
    columnas_esperadas = [
        "Voucher No", "Fecha", "Account", "Activity ID", "Payee / Vendor", 
        "Invoice No", "Description", "Payment / Ref", 
        "Budget Line No", "Payment (Bs)", "Expenses (87%)", "VAT Tax (13%)", "Glosa Original"
    ]
    for col in columnas_esperadas:
        if col not in df_res.columns:
            df_res[col] = ""

    return df_res, df_budget_cat
