import io
import re
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st


# ============================================================
# OPTIMIZACIÓN Y LIMPIEZA VECTORIZADA
# ============================================================
def clean_str_series(series: pd.Series) -> pd.Series:
    """Limpia cadenas de forma vectorizada."""
    return series.fillna("").astype(str).str.strip()


def norm_series(series: pd.Series) -> pd.Series:
    """Normaliza texto removiendo acentos y espacios vectorialmente."""
    s = series.fillna("").astype(str).str.strip().str.lower()
    return (
        s.str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("utf-8")
        .str.replace(r"\s+", " ", regex=True)
    )


def fast_to_number(series: pd.Series) -> pd.Series:
    """Convierte números rápidamente de forma vectorizada."""
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float)

    s = series.fillna("").astype(str).str.strip()
    is_neg = s.str.startswith("(") & s.str.endswith(")")
    s = s.str.replace(r"[()\$]", "", regex=True)
    s = s.str.replace("Bs.", "", regex=False).str.replace("Bs", "", regex=False)
    s = s.str.replace(" ", "", regex=False)

    # Manejo básico de formato numérico latino/inglés
    s = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)

    res = pd.to_numeric(s, errors="coerce")
    res = np.where(is_neg, -res, res)
    return pd.Series(res, index=series.index)


# ============================================================
# DETECCIÓN RÁPIDA DE ESTRUCTURA EN HOJAS
# ============================================================
def detect_pivot_header(raw: pd.DataFrame):
    for r in range(min(len(raw), 25)):
        row = raw.iloc[r].tolist()
        month_cols = []
        for c, val in enumerate(row):
            if pd.notna(val):
                val_str = str(val).strip()
                if re.search(r"20\d{2}", val_str) or isinstance(
                    val, (pd.Timestamp, datetime)
                ):
                    month_cols.append(c)

        if len(month_cols) >= 2:
            return {"row": r, "month_cols": month_cols}
    return None


def parse_gasto_sheet(raw: pd.DataFrame, sheet_name: str):
    detected = detect_pivot_header(raw)
    if not detected:
        return None

    header_row = detected["row"]
    month_cols = detected["month_cols"]

    # Extraer metadatos rápida
    company = "GENERAL"
    currency = "USD"
    for r in range(min(6, len(raw))):
        txt = " ".join([str(v) for v in raw.iloc[r].dropna().tolist()]).lower()
        if "empresa" in txt:
            company = sheet_name
        if "ves" in txt or "bs" in txt:
            currency = "VES"

    row_label_col = 0
    for c in range(min(5, raw.shape[1])):
        val_norm = str(raw.iloc[header_row, c]).strip().lower()
        if "etiquetas" in val_norm or "fila" in val_norm:
            row_label_col = c
            break

    # Fechas vectorizadas
    month_map = {}
    for c in month_cols:
        val = raw.iloc[header_row, c]
        dt = pd.to_datetime(val, errors="coerce")
        if pd.notna(dt):
            month_map[c] = pd.Timestamp(dt).to_period("M").to_timestamp()

    if not month_map:
        return None

    data_df = raw.iloc[header_row + 1 :].copy()
    labels = clean_str_series(data_df.iloc[:, row_label_col])

    # Ignorar totales
    mask_totales = (
        labels.str.lower().isin(["", "nan", "total general", "total", "grand total"])
        | labels.str.startswith("Total")
    )
    data_df = data_df[~mask_totales].copy()
    labels = labels[~mask_totales]

    records = []
    current_segment = "GENERAL"

    for idx, (_, row) in enumerate(data_df.iterrows()):
        lbl = labels.iloc[idx]
        if not lbl:
            continue

        is_account = bool(re.match(r"^\d+(?:\.\d+){2,}", lbl))
        if not is_account:
            current_segment = lbl
            level = "SEGMENTO"
        else:
            level = "CUENTA"

        for c, month_dt in month_map.items():
            val = fast_to_number(pd.Series([row[c]])).iloc[0]
            if pd.notna(val) and abs(val) > 0.001:
                records.append({
                    "pestana": sheet_name,
                    "empresa": company,
                    "moneda": currency,
                    "vista": current_segment,
                    "segmento": current_segment,
                    "cuenta": lbl,
                    "nivel": level,
                    "mes": month_dt,
                    "monto": float(val),
                })

    return pd.DataFrame(records) if records else None


# ============================================================
# CÁLCULOS OPTIMIZADOS EN MEMORIA
# ============================================================
@st.cache_data(show_spinner=False, max_entries=5)
def process_excel_lightweight(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    all_frames = []
    diagnostics = []

    for name in xls.sheet_names:
        if name.lower() == "detalle":
            continue

        try:
            raw = pd.read_excel(xls, sheet_name=name, header=None)
            parsed = parse_gasto_sheet(raw, name)
            if parsed is not None and not parsed.empty:
                all_frames.append(parsed)
                diagnostics.append({"Pestaña": name, "Estatus": "Procesada", "Registros": len(parsed)})
            else:
                diagnostics.append({"Pestaña": name, "Estatus": "Sin datos válidos", "Registros": 0})
        except Exception as e:
            diagnostics.append({"Pestaña": name, "Estatus": f"Error: {str(e)}", "Registros": 0})

    gastos = pd.concat(all_frames, ignore_index=True) if all_frames else pd.DataFrame()
    return gastos, pd.DataFrame(diagnostics), xls.sheet_names


def calculate_variations_fast(gastos):
    if gastos.empty:
        return pd.DataFrame()

    keys = ["pestana", "empresa", "moneda", "vista", "segmento", "cuenta", "nivel"]
    base = gastos.groupby(keys + ["mes"], as_index=False)["monto"].sum()
    base = base.sort_values(keys + ["mes"])

    base["mes_anterior"] = base.groupby(keys)["monto"].shift(1)
    base["variacion_bs"] = base["monto"] - base["mes_anterior"].fillna(0)

    base["variacion_pct"] = np.where(
        base["mes_anterior"].abs() > 0.01,
        (base["variacion_bs"] / base["mes_anterior"].abs()) * 100,
        np.nan,
    )

    base["promedio_6m"] = base.groupby(keys)["monto"].transform(
        lambda s: s.shift(1).rolling(6, min_periods=1).mean()
    )
    base["std_6m"] = base.groupby(keys)["monto"].transform(
        lambda s: s.shift(1).rolling(6, min_periods=2).std()
    )

    base["z_score"] = np.where(
        base["std_6m"].fillna(0) > 0.01,
        (base["monto"] - base["promedio_6m"]) / base["std_6m"],
        np.nan,
    )

    base["cambio_signo"] = (
        base["mes_anterior"].notna()
        & (base["mes_anterior"].abs() > 0.01)
        & (base["monto"].abs() > 0.01)
        & (np.sign(base["mes_anterior"]) != np.sign(base["monto"]))
    )

    base["nuevo_saldo"] = (
        base["mes_anterior"].fillna(0).abs().le(0.01)
        & base["monto"].abs().gt(0.01)
    )

    return base


def apply_alerts(df, pct_thresh, amount_thresh, z_thresh):
    if df.empty:
        return df

    cond_amount = df["variacion_bs"].abs() >= amount_thresh
    cond_pct = df["variacion_pct"].abs() >= pct_thresh
    cond_z = df["z_score"].abs() >= z_thresh
    cond_special = df["cambio_signo"] | df["nuevo_saldo"]

    hits = cond_amount.astype(int) + cond_pct.astype(int) + cond_z.astype(int)

    conditions = [
        cond_special & (hits > 0),
        hits >= 2,
        hits == 1,
    ]
    choices = ["🔴 CRÍTICA", "🔴 CRÍTICA", "🟠 IMPORTANTE"]

    df["nivel_alerta"] = np.select(conditions, choices, default="🟢 NORMAL")
    df["impacto_abs"] = df["variacion_bs"].abs()
    return df.sort_values(["mes", "impacto_abs"], ascending=[False, False])


# ============================================================
# INTERFAZ Y MÓDULO PRINCIPAL
# ============================================================
def modulo_analizador_gastos(sucursal: str):
    st.title("📊 Laboratorio de Variaciones de Gastos")
    st.caption(f"📍 Sede: **{sucursal}** | Análisis de variaciones acelerado.")

    uploaded = st.file_uploader("📂 Cargar el Excel de Gastos Operativos", type=["xlsx", "xls"])

    if uploaded is None:
        st.info("Sube un archivo Excel para iniciar el procesamiento.")
        with st.expander("📖 Instructivo y Parámetros Recomendados"):
            st.markdown(
                """
                ### 💡 Guía Rápida:
                1. **Variación % para alertar:** Identifica cambios porcentuales significativos (Recomendado: 20% a 30%).
                2. **Variación mínima ($/Bs):** Descarta variaciones pequeñas que no impactan los estados financieros.
                3. **Z-score atípico:** Marca desvíos estadísticos sobre el comportamiento habitual de los últimos 6 meses.
                """
            )
        st.stop()

    file_bytes = uploaded.getvalue()

    with st.spinner("Procesando datos en milisegundos..."):
        gastos, diagnostics, sheet_names = process_excel_lightweight(file_bytes)
        variations = calculate_variations_fast(gastos)

    # Parámetros Principales
    with st.expander("⚙️ Parámetros del Analizador de Alertas", expanded=True):
        col1, col2, col3, col4 = st.columns(4)
        pct_threshold = col1.number_input("Variación % min", min_value=1.0, value=30.0, step=5.0)
        amount_threshold = col2.number_input("Monto min ($/Bs)", min_value=0.0, value=5000.0, step=1000.0)
        z_threshold = col3.number_input("Z-Score min", min_value=1.0, value=2.0, step=0.5)
        exclude_nonexpense = col4.checkbox("Excluir Ingresos/Merma", value=False)

    if exclude_nonexpense and not variations.empty:
        variations = variations[
            ~variations["cuenta"].str.lower().str.contains("ingreso|merma", regex=True)
        ]

    variations = apply_alerts(variations, pct_threshold, amount_threshold, z_threshold)

    # Tabs
    tab_dash, tab_ranking, tab_detalle, tab_diag = st.tabs(
        ["🏠 Dashboard", "🚨 Ranking Alertas", "🔍 Detalle Transaccional", "🛠 Diagnóstico"]
    )

    with tab_dash:
        if variations.empty:
            st.warning("No se hallaron variaciones.")
        else:
            m_latest = variations["mes"].max()
            curr = variations[variations["mes"] == m_latest]

            c1, c2, c3 = st.columns(3)
            c1.metric("Período Reciente", m_latest.strftime("%Y-%m"))
            c2.metric("Alertas Críticas", len(curr[curr["nivel_alerta"] == "🔴 CRÍTICA"]))
            c3.metric("Alertas Importantes", len(curr[curr["nivel_alerta"] == "🟠 IMPORTANTE"]))

            st.dataframe(
                curr[["nivel_alerta", "pestana", "segmento", "cuenta", "monto", "mes_anterior", "variacion_bs", "variacion_pct"]],
                use_container_width=True,
                height=400,
            )

    with tab_ranking:
        st.subheader("Alertas por Nivel de Severidad")
        alerts_only = variations[variations["nivel_alerta"].isin(["🔴 CRÍTICA", "🟠 IMPORTANTE"])]
        st.dataframe(alerts_only, use_container_width=True, height=500)

    with tab_detalle:
        st.subheader("Carga perezosa de la Hoja Detalle")
        if "Detalle" in sheet_names:
            if st.button("📥 Parsear Hoja Detalle ahora"):
                with st.spinner("Procesando transacciones..."):
                    df_det = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Detalle")
                    st.success(f"Se cargaron {len(df_det):,} transacciones.")
                    st.dataframe(df_det.head(100), use_container_width=True)
        else:
            st.info("El archivo actual no contiene la hoja 'Detalle'.")

    with tab_diag:
        st.dataframe(diagnostics, use_container_width=True)
