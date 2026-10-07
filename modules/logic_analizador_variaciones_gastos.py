import io
import re
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st


# ============================================================
# NORMALIZACIÓN / PARSING
# ============================================================
def norm(x):
    if pd.isna(x):
        return ""
    s = str(x).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s)


def clean_label(x):
    if pd.isna(x):
        return ""
    return re.sub(r"\s+", " ", str(x).strip())


def is_month_label(x):
    if pd.isna(x):
        return False
    if isinstance(x, (pd.Timestamp, datetime)):
        return True

    s = str(x).strip()
    if re.match(r"^20\d{2}[-/]\d{1,2}$", s):
        return True
    if re.match(r"^20\d{2}[-/]\d{1,2}[-/]\d{1,2}$", s):
        return True
    if re.match(r"^[A-Za-zÁÉÍÓÚáéíóú]{3,12}[-/]\d{2,4}$", s):
        return True

    try:
        dt = pd.to_datetime(s, errors="coerce")
        return not pd.isna(dt)
    except Exception:
        return False


def to_month(x):
    if pd.isna(x):
        return pd.NaT
    if isinstance(x, (pd.Timestamp, datetime)):
        return pd.Timestamp(x).to_period("M").to_timestamp()

    s = str(x).strip()
    m = re.match(r"^(20\d{2})[-/](\d{1,2})(?:[-/]\d{1,2})?$", s)
    if m:
        try:
            return pd.Timestamp(int(m.group(1)), int(m.group(2)), 1)
        except Exception:
            pass

    try:
        dt = pd.to_datetime(s, errors="coerce")
        if not pd.isna(dt):
            return pd.Timestamp(dt).to_period("M").to_timestamp()
    except Exception:
        pass
    return pd.NaT


def to_number(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)

    s = str(x).strip()
    if not s:
        return np.nan
    negative = s.startswith("(") and s.endswith(")")
    s = s.replace("(", "").replace(")", "")
    s = s.replace("$", "").replace("Bs.", "").replace("Bs", "")
    s = s.replace(" ", "")

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        tail = s.rsplit(",", 1)[-1]
        if len(tail) in (1, 2):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")

    try:
        n = float(s)
        return -n if negative else n
    except Exception:
        return np.nan


def first_nonempty(values):
    for v in values:
        if not pd.isna(v) and str(v).strip():
            return str(v).strip()
    return ""


# ============================================================
# DETECCIÓN DE PESTAÑAS TIPO TABLA DINÁMICA
# ============================================================
def detect_pivot_header(raw):
    best = None
    for r in range(min(len(raw), 30)):
        row = raw.iloc[r].tolist()
        month_cols = [c for c, v in enumerate(row) if is_month_label(v)]
        if len(month_cols) < 3:
            continue

        has_row_label = any(
            norm(v) == "etiquetas de fila"
            for v in row[:8]
            if not pd.isna(v)
        )

        candidate = {
            "row": r,
            "month_cols": month_cols,
            "has_row_label": has_row_label,
        }

        if best is None:
            best = candidate
        elif has_row_label and not best["has_row_label"]:
            best = candidate
        elif len(month_cols) > len(best["month_cols"]):
            best = candidate
    return best


def sheet_metadata(raw):
    company = ""
    currency = ""

    for r in range(min(8, len(raw))):
        row = raw.iloc[r].tolist()
        for c, value in enumerate(row[:8]):
            if norm(value) == "empresa":
                company = first_nonempty(row[c + 1 : c + 4])
                break

    for r in range(min(8, len(raw))):
        for value in raw.iloc[r].tolist()[:8]:
            text = "" if pd.isna(value) else str(value)
            if "suma de" in norm(text):
                m = re.search(r"suma de\s+(.+)", norm(text))
                if m:
                    currency = m.group(1).strip().upper()
    return {"empresa": company, "moneda": currency or "USD"}


def is_account_line(label):
    if not label:
        return False
    return bool(re.match(r"^\d+(?:\.\d+){3,}\s*-\s*", str(label).strip()))


def is_total_label(label):
    return norm(label) in {"", "nan", "total general", "total", "grand total"}


def parse_gasto_sheet(raw, sheet_name):
    detected = detect_pivot_header(raw)
    if not detected:
        return None

    header_row = detected["row"]
    month_cols = detected["month_cols"]
    if len(month_cols) < 2:
        return None

    meta = sheet_metadata(raw)
    row_label_col = 0

    if norm(raw.iloc[header_row, 0]) != "etiquetas de fila":
        for c in range(min(8, raw.shape[1])):
            if norm(raw.iloc[header_row, c]) == "etiquetas de fila":
                row_label_col = c
                break

    month_map = {}
    for c in month_cols:
        month = to_month(raw.iloc[header_row, c])
        if not pd.isna(month):
            month_map[c] = month

    if not month_map:
        return None

    records = []
    current_view = ""
    current_segment = ""

    for r in range(header_row + 1, len(raw)):
        label = clean_label(raw.iloc[r, row_label_col])
        if is_total_label(label):
            continue

        values = {}
        for c, month in month_map.items():
            n = to_number(raw.iloc[r, c])
            values[month] = 0.0 if pd.isna(n) else n

        if not label or all(abs(v) < 1e-12 for v in values.values()):
            if not is_account_line(label):
                continue

        account_line = is_account_line(label)
        if not account_line:
            current_segment = label
            if norm(label).startswith("18.logistico"):
                current_view = label
            elif current_view == "" or "logist" not in norm(sheet_name):
                current_view = label
            level = "SEGMENTO"
        else:
            level = "CUENTA"

        if "logist" in norm(sheet_name):
            view = current_view or "18.Logistico"
        else:
            view = current_view or "General"

        segment = current_segment or label

        for month, amount in values.items():
            records.append({
                "pestana": sheet_name,
                "empresa": meta["empresa"],
                "moneda": meta["moneda"],
                "vista": view,
                "segmento": segment,
                "cuenta": label,
                "nivel": level,
                "mes": month,
                "monto": amount,
            })

    if not records:
        return None
    return pd.DataFrame(records)


# ============================================================
# HOJA DETALLE
# ============================================================
DETALLE_COLS = [
    "Empresa", "Periodo", "N2", "Cuenta_Nombre", "REFERENCIA", "Asiento",
    "CENTRO_COSTO", "USD", "VES", "USD_S1", "OFICINA", "SUB-GRUPO",
    "DEPARTAMENTO", "CONCATENADO", "N2 2.0", "nombre",
]


def prepare_detail(raw):
    if raw is None or raw.empty:
        return None, {"valido": False, "motivo": "Hoja vacía"}

    raw = raw.copy()
    raw.columns = [str(c).strip() for c in raw.iloc[0].tolist()]
    detail = raw.iloc[1:].copy()

    wanted = [c for c in DETALLE_COLS if c in detail.columns]
    detail = detail[wanted].copy()

    required = {"Empresa", "Periodo", "Cuenta_Nombre", "USD"}
    if not required.issubset(set(detail.columns)):
        return None, {
            "valido": False,
            "motivo": "No se encontraron las columnas mínimas de Detalle",
            "columnas": list(detail.columns),
        }

    detail["Periodo"] = detail["Periodo"].apply(to_month)
    detail["Empresa"] = detail["Empresa"].fillna("").astype(str).str.strip()

    for col in [
        "Cuenta_Nombre", "REFERENCIA", "Asiento", "CENTRO_COSTO", "OFICINA",
        "SUB-GRUPO", "DEPARTAMENTO", "CONCATENADO", "N2", "N2 2.0", "nombre",
    ]:
        if col in detail.columns:
            detail[col] = detail[col].fillna("").astype(str).str.strip()

    for col in ["USD", "VES", "USD_S1"]:
        if col in detail.columns:
            detail[col] = detail[col].apply(to_number)

    detail = detail.dropna(subset=["Periodo"]).copy()
    detail["_empresa"] = detail["Empresa"].map(norm)
    detail["_cuenta"] = detail["Cuenta_Nombre"].map(norm)
    for c in ["N2", "N2 2.0", "CONCATENADO", "nombre"]:
        if c in detail.columns:
            detail[f"_{c}"] = detail[c].map(norm)
    detail["_mes"] = detail["Periodo"]

    return detail, {
        "valido": True,
        "filas": len(detail),
        "columnas": list(detail.columns),
    }


# ============================================================
# CARGA CACHÉ (CORREGIDA PARA EVITAR OUT-OF-BOUNDS)
# ============================================================
@st.cache_data(show_spinner=False)
def read_workbook(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    raw_sheets = {}

    for name in xls.sheet_names:
        # Se elimina el 'usecols' estático para leer dinámicamente las columnas de cada pestaña
        raw_sheets[name] = pd.read_excel(
            xls,
            sheet_name=name,
            header=None,
        )

    return raw_sheets, xls.sheet_names


@st.cache_data(show_spinner=False)
def build_analysis(file_bytes):
    raw_sheets, sheet_names = read_workbook(file_bytes)
    all_frames = []
    diagnostics = []

    for name in sheet_names:
        if name == "Detalle":
            continue

        try:
            parsed = parse_gasto_sheet(raw_sheets[name], name)
            if parsed is None or parsed.empty:
                diagnostics.append({
                    "Pestaña": name,
                    "Analizada": "No",
                    "Registros": 0,
                    "Empresa": "",
                    "Moneda": "",
                    "Motivo": "No se detectó estructura mensual",
                })
                continue

            all_frames.append(parsed)
            diagnostics.append({
                "Pestaña": name,
                "Analizada": "Sí",
                "Registros": len(parsed),
                "Empresa": parsed["empresa"].iloc[0],
                "Moneda": parsed["moneda"].iloc[0],
                "Motivo": "",
            })
        except Exception as exc:
            diagnostics.append({
                "Pestaña": name,
                "Analizada": "Error",
                "Registros": 0,
                "Empresa": "",
                "Moneda": "",
                "Motivo": str(exc),
            })

    gastos = pd.concat(all_frames, ignore_index=True) if all_frames else pd.DataFrame()

    detail = None
    detail_status = {}
    if "Detalle" in raw_sheets:
        try:
            detail, detail_status = prepare_detail(raw_sheets["Detalle"])
        except Exception as exc:
            detail_status = {"valido": False, "motivo": str(exc)}

    return gastos, detail, pd.DataFrame(diagnostics), detail_status


# ============================================================
# HISTÓRICO Y VARIACIONES
# ============================================================
def complete_month_history(df):
    if df.empty:
        return df

    keys_cols = [
        "pestana", "empresa", "moneda", "vista",
        "segmento", "cuenta", "nivel",
    ]

    result = []
    for keys, group in df.groupby(keys_cols, dropna=False):
        group = group.sort_values("mes")
        months = pd.date_range(
            start=group["mes"].min(),
            end=group["mes"].max(),
            freq="MS",
        )
        values = group.set_index("mes")["monto"].reindex(months, fill_value=0.0)
        out = values.rename("monto").reset_index().rename(columns={"index": "mes"})
        for col, value in zip(keys_cols, keys):
            out[col] = value
        result.append(out)

    return pd.concat(result, ignore_index=True)


def calculate_variations(gastos):
    if gastos.empty:
        return pd.DataFrame()

    base = (
        gastos.groupby(
            [
                "pestana", "empresa", "moneda", "vista", "segmento",
                "cuenta", "nivel", "mes",
            ],
            dropna=False,
            as_index=False,
        )["monto"].sum()
    )

    base = complete_month_history(base)
    base = base.sort_values(
        ["pestana", "vista", "segmento", "cuenta", "nivel", "mes"]
    )

    group = [
        "pestana", "empresa", "moneda", "vista",
        "segmento", "cuenta", "nivel",
    ]

    base["mes_anterior"] = base.groupby(group)["monto"].shift(1)
    base["variacion_bs"] = base["monto"] - base["mes_anterior"].fillna(0)

    base["variacion_pct"] = np.where(
        base["mes_anterior"].abs() > 0.01,
        base["variacion_bs"] / base["mes_anterior"].abs() * 100,
        np.nan,
    )

    base["promedio_3m"] = base.groupby(group)["monto"].transform(
        lambda s: s.shift(1).rolling(3, min_periods=1).mean()
    )
    base["promedio_6m"] = base.groupby(group)["monto"].transform(
        lambda s: s.shift(1).rolling(6, min_periods=1).mean()
    )
    base["mediana_6m"] = base.groupby(group)["monto"].transform(
        lambda s: s.shift(1).rolling(6, min_periods=2).median()
    )
    base["std_6m"] = base.groupby(group)["monto"].transform(
        lambda s: s.shift(1).rolling(6, min_periods=2).std()
    )

    base["z_score"] = np.where(
        base["std_6m"].abs() > 0.01,
        (base["monto"] - base["mediana_6m"]) / base["std_6m"],
        np.nan,
    )

    base["cambio_signo"] = (
        base["mes_anterior"].notna()
        & (base["mes_anterior"].abs() > 0.01)
        & (base["monto"].abs() > 0.01)
        & (np.sign(base["mes_anterior"]) != np.sign(base["monto"]))
    )

    base["nuevo_saldo"] = (
        base["mes_anterior"].abs().fillna(0).le(0.01)
        & base["monto"].abs().gt(0.01)
    )

    total_keys = ["pestana", "empresa", "moneda", "vista", "mes"]
    totals = (
        base.groupby(total_keys, as_index=False)["monto"]
        .sum()
        .rename(columns={"monto": "total_pestana_mes"})
    )
    base = base.merge(totals, on=total_keys, how="left")

    base["peso_en_total_pct"] = np.where(
        base["total_pestana_mes"].abs() > 0.01,
        base["monto"].abs() / base["total_pestana_mes"].abs() * 100,
        np.nan,
    )

    delta_total = (
        base.groupby(total_keys, as_index=False)["variacion_bs"]
        .agg(suma_abs_variaciones=lambda s: s.abs().sum())
    )
    base = base.merge(delta_total, on=total_keys, how="left")
    base["participacion_variacion_pct"] = np.where(
        base["suma_abs_variaciones"] > 0.01,
        base["variacion_bs"].abs() / base["suma_abs_variaciones"] * 100,
        0,
    )

    base["direccion"] = np.select(
        [
            base["cambio_signo"],
            base["variacion_bs"].gt(0.01),
            base["variacion_bs"].lt(-0.01),
        ],
        ["CAMBIO DE SIGNO", "AUMENTO", "DISMINUCIÓN"],
        default="SIN CAMBIO",
    )

    return base


def classify_alert(row, pct_threshold, amount_threshold, z_threshold):
    amount_hit = abs(float(row["variacion_bs"] or 0)) >= amount_threshold
    pct = row["variacion_pct"]
    pct_hit = not pd.isna(pct) and abs(float(pct)) >= pct_threshold
    z_hit = not pd.isna(row["z_score"]) and abs(float(row["z_score"])) >= z_threshold
    sign_hit = bool(row["cambio_signo"])
    new_hit = bool(row["nuevo_saldo"])

    hits = sum([amount_hit, pct_hit, z_hit, sign_hit, new_hit])

    if sign_hit or new_hit:
        return "🔴 CRÍTICA" if (amount_hit or pct_hit or z_hit) else "🟠 IMPORTANTE"
    if hits >= 2:
        return "🔴 CRÍTICA"
    if hits == 1:
        return "🟠 IMPORTANTE"
    return "🟢 NORMAL"


def add_alert_classification(variations, pct_threshold, amount_threshold, z_threshold):
    if variations.empty:
        return variations

    x = variations.copy()
    x["nivel_alerta"] = x.apply(
        lambda row: classify_alert(row, pct_threshold, amount_threshold, z_threshold),
        axis=1,
    )
    x["impacto_abs"] = x["variacion_bs"].abs()
    priority = {"🔴 CRÍTICA": 0, "🟠 IMPORTANTE": 1, "🟢 NORMAL": 2}
    x["prioridad"] = x["nivel_alerta"].map(priority)
    return x.sort_values(
        ["mes", "prioridad", "impacto_abs", "participacion_variacion_pct"],
        ascending=[False, True, False, False],
    )


# ============================================================
# ENLACE PESTAÑA -> DETALLE
# ============================================================
def detail_matches_alert(alert, detail):
    if detail is None or detail.empty:
        return pd.DataFrame()

    d = detail.copy()
    subset = d[
        d["_empresa"].eq(norm(alert["empresa"]))
        & d["_mes"].eq(alert["mes"])
    ].copy()

    if subset.empty:
        return subset

    account = str(alert["cuenta"]).strip()
    segment = str(alert["segmento"]).strip()
    view = str(alert["vista"]).strip()

    if alert["nivel"] == "CUENTA":
        exact = subset[subset["_cuenta"].eq(norm(account))].copy()
        if not exact.empty:
            return exact

        code = account.split("-", 1)[0].strip()
        if code:
            exact = subset[
                subset["_cuenta"].str.startswith(norm(code), na=False)
            ].copy()
            if not exact.empty:
                return exact

    segment_norm = norm(segment)
    view_norm = norm(view)
    n2_match = pd.Series(False, index=subset.index)

    if "_N2" in subset.columns:
        n2_match |= subset["_N2"].eq(segment_norm)
    if "_N2 2.0" in subset.columns:
        n2_match |= subset["_N2 2.0"].eq(segment_norm)

    if "logist" in norm(alert["pestana"]):
        if "_N2 2.0" in subset.columns and view_norm == "18.logistico":
            logistic_mask = subset["_N2 2.0"].eq(view_norm)
            if "_N2" in subset.columns:
                match = logistic_mask & subset["_N2"].eq(segment_norm)
                if match.any():
                    return subset[match].copy()

        if n2_match.any():
            return subset[n2_match].copy()

    if n2_match.any():
        return subset[n2_match].copy()

    candidates = pd.Series(False, index=subset.index)
    for col in ["_N2", "_N2 2.0", "_CONCATENADO", "_nombre"]:
        if col in subset.columns:
            candidates |= subset[col].str.contains(
                re.escape(segment_norm),
                case=False,
                na=False,
            )
    return subset[candidates].copy() if candidates.any() else pd.DataFrame()


def transaction_driver_table(detail_rows, amount_col):
    if detail_rows is None or detail_rows.empty:
        return pd.DataFrame()
    if amount_col not in detail_rows.columns:
        amount_col = "USD"

    d = detail_rows.copy()
    d["_impacto_abs"] = d[amount_col].abs()

    group_cols = [
        c for c in [
            "Cuenta_Nombre", "REFERENCIA", "Asiento",
            "CENTRO_COSTO", "DEPARTAMENTO", "nombre",
        ] if c in d.columns
    ]
    if not group_cols:
        return pd.DataFrame()

    out = (
        d.groupby(group_cols, dropna=False, as_index=False)
        .agg(
            movimientos=(amount_col, "count"),
            importe=(amount_col, "sum"),
            impacto_abs=("_impacto_abs", "sum"),
        )
        .sort_values("impacto_abs", ascending=False)
        .head(15)
    )
    return out


def fmt(df, cols):
    return df.style.format(
        {c: "{:,.2f}" for c in cols if c in df.columns},
        na_rep="-",
    )


# ============================================================
# FUNCIÓN PRINCIPAL DEL MÓDULO EXPORTABLE
# ============================================================
def modulo_analizador_gastos(sucursal: str):
    st.title("📊 Laboratorio de Variaciones de Gastos")
    st.caption(
        f"📍 Operando en: **{sucursal}** | "
        "Análisis mensual por pestaña, segmento, cuenta y transacción histórica. "
        "El archivo de gastos se analiza junto con la hoja Detalle."
    )

    uploaded = st.file_uploader(
        "📂 Cargar el Excel de Gastos Operativos",
        type=["xlsx", "xls"],
    )

    if uploaded is None:
        st.info("Carga el archivo **Gastos Operativos Mayoreo VE** para comenzar.")
        with st.expander("Estructura del archivo que espera esta versión"):
            st.markdown(
                """
                Esta versión está ajustada al archivo entregado: las pestañas de gastos
                tienen una fila con **Etiquetas de fila** y meses como `2026-09`,
                `2026-08`, `2026-07`; la columna A contiene segmentos y, en las
                pestañas `*_Categoria_Logist`, aparecen cuentas detalladas como
                `7.1.1.01.1.001-Nómina` debajo de segmentos.

                La hoja **Detalle** utiliza `Empresa`, `Periodo`, `N2`, `Cuenta_Nombre`,
                `REFERENCIA`, `Asiento`, `CENTRO_COSTO`, `USD`, `VES`, `USD_S1`,
                `DEPARTAMENTO`, `CONCATENADO`, `N2 2.0` y `nombre` para llegar desde
                una alerta hasta las transacciones que la explican.
                """
            )
        st.stop()

    file_bytes = uploaded.getvalue()

    with st.spinner("Leyendo y estructurando el archivo..."):
        try:
            raw_sheets, sheet_names = read_workbook(file_bytes)
            gastos, detail, diagnostics, detail_status = build_analysis(file_bytes)
        except Exception as exc:
            st.error(f"No fue posible procesar el archivo: {exc}")
            st.stop()

    # Parámetros en barra lateral
    st.sidebar.header("⚙️ Parámetros del Analizador")

    pct_threshold = st.sidebar.number_input(
        "Variación % para alertar",
        min_value=1.0,
        max_value=500.0,
        value=30.0,
        step=5.0,
    )
    amount_threshold = st.sidebar.number_input(
        "Variación mínima (moneda de la pestaña)",
        min_value=0.0,
        value=10000.0,
        step=5000.0,
    )
    z_threshold = st.sidebar.number_input(
        "Z-score atípico",
        min_value=1.0,
        max_value=8.0,
        value=2.0,
        step=0.5,
    )
    exclude_nonexpense = st.sidebar.checkbox(
        "Excluir Ingresos y Merma",
        value=False,
    )

    variations = calculate_variations(gastos)

    if exclude_nonexpense and not variations.empty:
        keep = ~variations["cuenta"].map(norm).str.contains(
            r"ingresos|merma", regex=True, na=False
        )
        variations = variations[keep].copy()

    variations = add_alert_classification(
        variations,
        pct_threshold,
        amount_threshold,
        z_threshold,
    )

    # Filtros
    st.sidebar.divider()
    st.sidebar.header("🔎 Filtros del Analizador")

    analysis_sheets = sorted(gastos["pestana"].unique()) if not gastos.empty else []
    selected_sheets = st.sidebar.multiselect(
        "Pestañas",
        analysis_sheets,
        default=analysis_sheets,
    )

    available_months = (
        sorted(gastos.loc[gastos["pestana"].isin(selected_sheets), "mes"].dropna().unique())
        if not gastos.empty else []
    )

    if available_months:
        min_date = pd.Timestamp(min(available_months))
        max_date = pd.Timestamp(max(available_months))
        date_range = st.sidebar.date_input(
            "Período",
            value=(min_date.date(), max_date.date()),
        )
        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_date = pd.Timestamp(date_range[0]).to_period("M").to_timestamp()
            end_date = pd.Timestamp(date_range[1]).to_period("M").to_timestamp()
        else:
            start_date = min_date
            end_date = pd.Timestamp(date_range).to_period("M").to_timestamp()
    else:
        start_date = pd.Timestamp("1900-01-01")
        end_date = pd.Timestamp("2100-01-01")

    filtered = (
        variations[
            variations["pestana"].isin(selected_sheets)
            & variations["mes"].between(start_date, end_date)
        ].copy()
        if not variations.empty else pd.DataFrame()
    )

    # Tabs
    tab_dashboard, tab_pestanas, tab_ranking, tab_detalle, tab_diag = st.tabs(
        [
            "🏠 Dashboard",
            "📑 Pestañas",
            "🚨 Ranking de variaciones",
            "🔍 Detalle / causas",
            "🛠 Diagnóstico",
        ]
    )

    # Dashboard
    with tab_dashboard:
        st.subheader("Resumen ejecutivo")

        if filtered.empty:
            st.warning("No hay datos en el período o filtros seleccionados.")
        else:
            latest_month = filtered["mes"].max()
            latest = filtered[filtered["mes"] == latest_month].copy()

            critical = (latest["nivel_alerta"] == "🔴 CRÍTICA").sum()
            important = (latest["nivel_alerta"] == "🟠 IMPORTANTE").sum()
            total_latest = latest["monto"].sum()
            total_delta = latest["variacion_bs"].sum()

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Último mes", latest_month.strftime("%Y-%m"))
            c2.metric("Gasto analizado", f"{total_latest:,.2f}")
            c3.metric("Variación neta", f"{total_delta:,.2f}")
            c4.metric("Alertas críticas", f"{critical:,}")
            c5.metric("Alertas importantes", f"{important:,}")

            st.markdown("### 🔥 Top 20 variaciones que merecen revisión")
            top20 = latest.sort_values("impacto_abs", ascending=False).head(20).copy()
            dashboard_cols = [
                "nivel_alerta", "pestana", "empresa", "vista", "segmento",
                "cuenta", "mes", "monto", "mes_anterior", "variacion_bs",
                "variacion_pct", "peso_en_total_pct",
                "participacion_variacion_pct", "direccion",
            ]
            dashboard_cols = [c for c in dashboard_cols if c in top20.columns]
            st.dataframe(
                fmt(
                    top20[dashboard_cols],
                    [
                        "monto", "mes_anterior", "variacion_bs", "variacion_pct",
                        "peso_en_total_pct", "participacion_variacion_pct",
                    ],
                ),
                use_container_width=True,
                height=560,
            )

            st.markdown("### 📈 Evolución del gasto total")
            monthly = (
                filtered.groupby("mes", as_index=False)["monto"]
                .sum()
                .sort_values("mes")
            )
            st.line_chart(monthly.set_index("mes")[["monto"]])

            st.markdown("### 🧩 Segmentos que más explican el cambio")
            seg = (
                latest.groupby(["pestana", "vista", "segmento"], as_index=False)
                .agg(
                    gasto=("monto", "sum"),
                    variacion=("variacion_bs", "sum"),
                    impacto=("variacion_bs", lambda s: s.abs().sum()),
                )
                .sort_values("impacto", ascending=False)
                .head(15)
            )
            st.dataframe(
                fmt(seg, ["gasto", "variacion", "impacto"]),
                use_container_width=True,
            )

    # Análisis por Pestaña
    with tab_pestanas:
        st.subheader("📑 Variación mes a mes por pestaña")

        if not selected_sheets:
            st.info("Seleccione al menos una pestaña.")
        else:
            selected_sheet = st.selectbox("Pestaña", selected_sheets)
            sheet_data = filtered[filtered["pestana"] == selected_sheet].copy()

            if sheet_data.empty:
                st.warning("No hay datos para la pestaña seleccionada.")
            else:
                monthly = (
                    sheet_data.groupby("mes", as_index=False)
                    .agg(
                        gasto=("monto", "sum"),
                        variacion=("variacion_bs", "sum"),
                        registros=("cuenta", "count"),
                    )
                    .sort_values("mes")
                )

                st.markdown("### Evolución del gasto")
                st.line_chart(monthly.set_index("mes")[["gasto"]])

                st.markdown("### Resumen mensual")
                st.dataframe(
                    fmt(monthly, ["gasto", "variacion"]),
                    use_container_width=True,
                )

                segments = sorted(sheet_data["segmento"].dropna().astype(str).unique())
                selected_segment = st.selectbox(
                    "Segmento / grupo",
                    ["TODOS"] + segments,
                )

                entity = (
                    sheet_data
                    if selected_segment == "TODOS"
                    else sheet_data[sheet_data["segmento"].astype(str) == selected_segment]
                ).copy()

                accounts = sorted(entity["cuenta"].dropna().astype(str).unique())
                selected_account = "TODAS"
                if len(accounts) > 1:
                    selected_account = st.selectbox(
                        "Cuenta específica",
                        ["TODAS"] + accounts,
                    )

                if selected_account != "TODAS":
                    entity = entity[entity["cuenta"].astype(str) == selected_account].copy()

                st.markdown("### Tendencia seleccionada")
                trend = entity.groupby("mes")["monto"].sum().sort_index()
                st.line_chart(trend)

                detail_cols = [
                    "mes", "vista", "segmento", "cuenta", "nivel",
                    "monto", "mes_anterior", "variacion_bs", "variacion_pct",
                    "promedio_3m", "promedio_6m", "z_score",
                    "nivel_alerta", "direccion",
                ]
                st.dataframe(
                    fmt(
                        entity[[c for c in detail_cols if c in entity.columns]],
                        [
                            "monto", "mes_anterior", "variacion_bs", "variacion_pct",
                            "promedio_3m", "promedio_6m", "z_score",
                        ],
                    ),
                    use_container_width=True,
                    height=560,
                )

    # Ranking
    with tab_ranking:
        st.subheader("🚨 Ranking de variaciones importantes")

        if filtered.empty:
            st.info("No existen variaciones para los filtros seleccionados.")
        else:
            ranking = filtered[
                filtered["nivel_alerta"].isin(["🔴 CRÍTICA", "🟠 IMPORTANTE"])
            ].copy()

            direction = st.multiselect(
                "Tipo de movimiento",
                ["AUMENTO", "DISMINUCIÓN", "CAMBIO DE SIGNO"],
                default=["AUMENTO", "DISMINUCIÓN", "CAMBIO DE SIGNO"],
            )
            if direction:
                ranking = ranking[ranking["direccion"].isin(direction)]

            sort_choice = st.radio(
                "Ordenar por",
                [
                    "Impacto económico",
                    "Variación %",
                    "Z-score",
                    "Participación del cambio",
                ],
                horizontal=True,
            )
            sort_map = {
                "Impacto económico": "impacto_abs",
                "Variación %": "variacion_pct",
                "Z-score": "z_score",
                "Participación del cambio": "participacion_variacion_pct",
            }
            sort_col = sort_map[sort_choice]
            ranking = ranking.sort_values(
                sort_col,
                key=lambda s: s.abs(),
                ascending=False,
            )

            ranking_cols = [
                "nivel_alerta", "mes", "pestana", "empresa", "vista",
                "segmento", "cuenta", "monto", "mes_anterior", "variacion_bs",
                "variacion_pct", "promedio_3m", "promedio_6m", "z_score",
                "peso_en_total_pct", "participacion_variacion_pct",
                "direccion", "cambio_signo", "nuevo_saldo",
            ]
            st.dataframe(
                fmt(
                    ranking[[c for c in ranking_cols if c in ranking.columns]],
                    [
                        "monto", "mes_anterior", "variacion_bs", "variacion_pct",
                        "promedio_3m", "promedio_6m", "z_score",
                        "peso_en_total_pct", "participacion_variacion_pct",
                    ],
                ),
                use_container_width=True,
                height=650,
            )

    # Detalle / Causas
    with tab_detalle:
        st.subheader("🔍 Desde la alerta hasta la transacción")

        if detail is None or not detail_status.get("valido", False):
            st.warning(
                "No se pudo preparar la hoja Detalle. "
                "El análisis mensual seguirá funcionando, pero no el nivel transaccional."
            )
            st.write(detail_status)
        elif filtered.empty:
            st.info("No existen registros para los filtros actuales.")
        else:
            alerts = filtered[
                filtered["nivel_alerta"].isin(["🔴 CRÍTICA", "🟠 IMPORTANTE"])
            ].copy().reset_index(drop=True)

            if alerts.empty:
                st.success("No hay alertas importantes en el período seleccionado.")
            else:
                def label_alert(i):
                    r = alerts.iloc[i]
                    pct = f"{r['variacion_pct']:,.2f}%" if not pd.isna(r['variacion_pct']) else "N/D"
                    return (
                        f"{r['nivel_alerta']} | {r['pestana']} | {r['segmento']} | "
                        f"{r['cuenta']} | {pd.Timestamp(r['mes']).strftime('%Y-%m')} | "
                        f"Δ {r['variacion_bs']:,.2f} | {pct}"
                    )

                selected_idx = st.selectbox(
                    "Seleccione una alerta",
                    range(len(alerts)),
                    format_func=label_alert,
                )
                alert = alerts.iloc[selected_idx]

                k1, k2, k3, k4, k5 = st.columns(5)
                k1.metric("Saldo actual", f"{alert['monto']:,.2f}")
                k2.metric("Mes anterior", f"{alert['mes_anterior']:,.2f}")
                k3.metric("Variación", f"{alert['variacion_bs']:,.2f}")
                k4.metric(
                    "Variación %",
                    f"{alert['variacion_pct']:,.2f}%" if not pd.isna(alert['variacion_pct']) else "N/D",
                )
                k5.metric(
                    "Z-score",
                    f"{alert['z_score']:,.2f}" if not pd.isna(alert['z_score']) else "N/D",
                )

                reasons = []
                if abs(alert["variacion_bs"]) >= amount_threshold:
                    reasons.append(f"supera el umbral monetario de {amount_threshold:,.2f}")
                if not pd.isna(alert["variacion_pct"]) and abs(alert["variacion_pct"]) >= pct_threshold:
                    reasons.append(f"la variación porcentual es {abs(alert['variacion_pct']):,.2f}%")
                if not pd.isna(alert["z_score"]) and abs(alert["z_score"]) >= z_threshold:
                    reasons.append(f"presenta comportamiento atípico (Z-score {abs(alert['z_score']):,.2f})")
                if alert["cambio_signo"]:
                    reasons.append("existe cambio de signo")
                if alert["nuevo_saldo"]:
                    reasons.append("es un saldo nuevo respecto del mes anterior")

                st.info(
                    "La alerta se explica por: " + "; ".join(reasons) + "."
                    if reasons
                    else "La fila no supera actualmente los parámetros configurados."
                )

                detail_rows = detail_matches_alert(alert, detail)

                amount_col = (
                    "USD_S1"
                    if norm(alert["moneda"]) == "usd_s1" and "USD_S1" in detail_rows.columns
                    else "USD"
                )

                st.markdown("### 🔬 Transacciones que explican la variación")

                if detail_rows.empty:
                    st.warning(
                        "No se encontraron transacciones con el mismo Empresa + Periodo + Segmento/Cuenta."
                    )
                else:
                    st.metric("Movimientos encontrados", f"{len(detail_rows):,}")

                    drivers = transaction_driver_table(detail_rows, amount_col)
                    if not drivers.empty:
                        st.markdown("#### Principales movimientos")
                        st.dataframe(
                            fmt(drivers, ["movimientos", "importe", "impacto_abs"]),
                            use_container_width=True,
                            height=360,
                        )

                    st.markdown("#### Transacciones")
                    display_cols = [
                        c for c in DETALLE_COLS if c in detail_rows.columns
                    ]
                    st.dataframe(
                        fmt(detail_rows[display_cols], ["USD", "VES", "USD_S1"]),
                        use_container_width=True,
                        height=550,
                    )

    # Diagnóstico
    with tab_diag:
        st.subheader("🛠 Diagnóstico del archivo")
        st.dataframe(diagnostics, use_container_width=True, height=420)

        st.markdown("### Hoja Detalle")
        st.write(detail_status)
        if detail is not None and not detail.empty:
            st.write(f"Registros históricos utilizables: {len(detail):,}")
            st.write(
                "Columnas utilizadas:",
                [c for c in DETALLE_COLS if c in detail.columns],
            )

    # Exportación
    st.divider()
    st.subheader("📥 Exportar análisis")

    if not filtered.empty:
        export_variations = filtered.copy()
        export_alerts = filtered[
            filtered["nivel_alerta"].isin(["🔴 CRÍTICA", "🟠 IMPORTANTE"])
        ].copy()
        latest_export = (
            filtered[filtered["mes"] == filtered["mes"].max()]
            .sort_values("impacto_abs", ascending=False)
            .head(20)
        )

        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            export_variations.to_excel(writer, sheet_name="Variaciones", index=False)
            export_alerts.to_excel(writer, sheet_name="Alertas", index=False)
            latest_export.to_excel(writer, sheet_name="Top20_Ultimo_Mes", index=False)
            diagnostics.to_excel(writer, sheet_name="Diagnostico", index=False)

            if detail is not None and not detail.empty and not export_alerts.empty:
                parts = []
                for _, alert in export_alerts.head(100).iterrows():
                    rows = detail_matches_alert(alert, detail)
                    if not rows.empty:
                        parts.append(rows)
                if parts:
                    pd.concat(parts, ignore_index=True).drop_duplicates().to_excel(
                        writer,
                        sheet_name="Detalle_Alertas",
                        index=False,
                    )

            for ws in writer.book.worksheets:
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                for col in ws.columns:
                    max_len = 0
                    letter = col[0].column_letter
                    for cell in col[:1000]:
                        try:
                            max_len = max(max_len, len(str(cell.value)))
                        except Exception:
                            pass
                    ws.column_dimensions[letter].width = min(max(10, max_len + 2), 35)

        out.seek(0)
        st.download_button(
            "📥 Descargar reporte Excel",
            data=out.getvalue(),
            file_name=f"Analisis_Variaciones_Gastos_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("No hay datos para exportar con los filtros actuales.")
