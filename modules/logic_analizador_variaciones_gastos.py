
import io
import re
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURACIÓN
# ============================================================
st.set_page_config(
    page_title="Analizador de Variaciones de Gastos",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Analizador de Variaciones de Gastos")
st.caption(
    "Análisis mensual por pestaña, segmento, cuenta y transacción histórica "
    "usando la hoja Detalle como fuente de auditoría."
)

TIPOS_HOJA_IGNORAR = {
    "detalle", "resumen", "resumen ejecutivo", "dashboard",
    "config", "configuracion", "configuración"
}

# ============================================================
# FUNCIONES GENERALES
# ============================================================
def normalizar_texto(x):
    if pd.isna(x):
        return ""
    x = str(x).strip().lower()
    x = unicodedata.normalize("NFKD", x)
    x = "".join(c for c in x if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", x)


def limpiar_nombre_columna(x):
    s = normalizar_texto(x)
    s = re.sub(r"[^a-z0-9]+", "_", s).strip("_")
    return s


def convertir_numero(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)

    s = str(x).strip()
    if not s:
        return np.nan

    # Soporta:
    # 1,234.56
    # 1.234,56
    # 1 234,56
    # (1.234,56)
    negativo = s.startswith("(") and s.endswith(")")
    s = s.replace("(", "").replace(")", "")
    s = s.replace("$", "").replace("Bs.", "").replace("Bs", "")
    s = s.replace(" ", "")

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        partes = s.split(",")
        if len(partes[-1]) in (1, 2):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        # Punto decimal o miles; se deja como está.
        pass

    try:
        n = float(s)
        return -n if negativo else n
    except Exception:
        return np.nan


def leer_excel(uploaded_file):
    data = uploaded_file.getvalue()
    nombre = uploaded_file.name.lower()

    if nombre.endswith(".xls"):
        try:
            return pd.ExcelFile(io.BytesIO(data), engine="xlrd")
        except Exception as e:
            raise RuntimeError(
                "Para archivos .xls instala xlrd: pip install xlrd"
            ) from e

    return pd.ExcelFile(io.BytesIO(data), engine="openpyxl")


def leer_hoja(xls, sheet_name):
    # Primero intenta con encabezado normal.
    try:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
        if df.shape[1] > 0:
            return df
    except Exception:
        pass

    return pd.read_excel(xls, sheet_name=sheet_name, header=None)


def es_mes(valor):
    """
    Reconoce:
      2026-09
      2026/09
      2026-09-01
      Sep-26
      2026-09-30
    """
    if pd.isna(valor):
        return False

    if isinstance(valor, (pd.Timestamp, datetime)):
        return True

    s = str(valor).strip()

    patrones = [
        r"^\d{4}[-/]\d{1,2}$",
        r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}$",
        r"^[A-Za-zÁÉÍÓÚáéíóú]{3,9}[-/]\d{2,4}$",
        r"^\d{1,2}[-/][A-Za-zÁÉÍÓÚáéíóú]{3,9}[-/]\d{2,4}$",
    ]
    return any(re.match(p, s) for p in patrones)


def convertir_mes(valor):
    if pd.isna(valor):
        return pd.NaT

    if isinstance(valor, (pd.Timestamp, datetime)):
        return pd.Timestamp(valor).to_period("M").to_timestamp()

    s = str(valor).strip()

    # yyyy-mm / yyyy/mm
    m = re.match(r"^(\d{4})[-/](\d{1,2})(?:[-/]\d{1,2})?$", s)
    if m:
        try:
            return pd.Timestamp(int(m.group(1)), int(m.group(2)), 1)
        except Exception:
            pass

    # Intento general
    try:
        dt = pd.to_datetime(s, errors="coerce", dayfirst=False)
        if not pd.isna(dt):
            return pd.Timestamp(dt).to_period("M").to_timestamp()
    except Exception:
        pass

    return pd.NaT


def detectar_columnas_mes(df):
    resultado = []

    for col in df.columns:
        # Detectar por nombre
        if es_mes(col):
            resultado.append(col)
            continue

        # Detectar si la columna tiene valores mensuales
        serie = df[col].dropna().head(20)
        if len(serie) > 0:
            cantidad = sum(es_mes(v) for v in serie)
            if cantidad >= max(2, int(len(serie) * 0.6)):
                resultado.append(col)

    return list(dict.fromkeys(resultado))


def detectar_columna(df, candidatos):
    mapa = {
        normalizar_texto(c): c
        for c in df.columns
    }

    for candidato in candidatos:
        nc = normalizar_texto(candidato)
        if nc in mapa:
            return mapa[nc]

    for col in df.columns:
        ncol = normalizar_texto(col)
        for candidato in candidatos:
            nc = normalizar_texto(candidato)
            if nc in ncol or ncol in nc:
                return col

    return None


def buscar_columna_por_patrones(df, patrones):
    for col in df.columns:
        n = normalizar_texto(col)
        if any(p in n for p in patrones):
            return col
    return None


# ============================================================
# DETECCIÓN DEL FORMATO DE LAS PESTAÑAS
# ============================================================
def detectar_estructura_pestana(df):
    meses = detectar_columnas_mes(df)

    if not meses:
        # También intenta detectar columnas que parezcan YYYY-MM
        meses = [
            c for c in df.columns
            if re.search(r"\b20\d{2}[-/]\d{1,2}\b", str(c))
        ]

    columna_cuenta = detectar_columna(
        df,
        ["cuenta", "codigo cuenta", "código cuenta", "account", "codigo"]
    )

    columna_descripcion = detectar_columna(
        df,
        ["etiquetas de fila", "descripcion", "descripción", "cuenta descripcion",
         "nombre cuenta", "cuenta", "segmento"]
    )

    return {
        "meses": meses,
        "cuenta": columna_cuenta,
        "descripcion": columna_descripcion,
    }


def convertir_pestana_a_largo(df, nombre_hoja):
    estructura = detectar_estructura_pestana(df)
    meses = estructura["meses"]

    if not meses:
        return None

    id_cols = [c for c in df.columns if c not in meses]

    # Conserva solo columnas de identificación útiles.
    if not id_cols:
        id_cols = []

    long = df.melt(
        id_vars=id_cols,
        value_vars=meses,
        var_name="mes_original",
        value_name="monto"
    )

    long["mes"] = long["mes_original"].apply(convertir_mes)
    long["monto"] = long["monto"].apply(convertir_numero)
    long["pestana"] = nombre_hoja

    # Selección de descripción / segmento.
    col_desc = estructura["descripcion"]
    col_cuenta = estructura["cuenta"]

    if col_desc and col_desc in long.columns:
        long["segmento"] = long[col_desc].astype(str).replace("nan", "").str.strip()
    elif col_cuenta and col_cuenta in long.columns:
        long["segmento"] = long[col_cuenta].astype(str).replace("nan", "").str.strip()
    else:
        # Primera columna no mensual como etiqueta.
        candidatos = [c for c in id_cols if long[c].dtype == "object"]
        if candidatos:
            long["segmento"] = long[candidatos[0]].astype(str).replace("nan", "").str.strip()
        else:
            long["segmento"] = nombre_hoja

    if col_cuenta and col_cuenta in long.columns:
        long["cuenta"] = long[col_cuenta].astype(str).replace("nan", "").str.strip()
    else:
        long["cuenta"] = long["segmento"]

    long["monto"] = long["monto"].fillna(0)

    # Elimina filas sin mes.
    long = long.dropna(subset=["mes"]).copy()

    return long


# ============================================================
# DETALLE HISTÓRICO
# ============================================================
def preparar_detalle(df):
    if df is None or df.empty:
        return None

    original = df.copy()
    original.columns = [str(c).strip() for c in original.columns]

    col_fecha = detectar_columna(
        original,
        ["fecha", "fecha contabilizacion", "fecha contable", "date"]
    )
    col_cuenta = detectar_columna(
        original,
        ["cuenta", "codigo cuenta", "código cuenta", "account"]
    )
    col_segmento = detectar_columna(
        original,
        ["segmento", "categoria", "categoría", "subcategoria",
         "subcategoría", "centro de costo", "centro costo", "rubro"]
    )
    col_descripcion = detectar_columna(
        original,
        ["descripcion", "descripción", "detalle", "concepto", "glosa", "nombre"]
    )
    col_monto = detectar_columna(
        original,
        ["monto", "importe", "valor", "debitos", "débitos", "gasto",
         "monto local", "debe", "saldo"]
    )

    # Si no encuentra monto, intenta detectar la columna numérica con mayor
    # proporción de valores convertibles.
    if col_monto is None:
        mejor = None
        mejor_score = 0
        for col in original.columns:
            serie = original[col].dropna()
            if len(serie) == 0:
                continue
            score = serie.map(convertir_numero).notna().mean()
            if score > mejor_score:
                mejor_score = score
                mejor = col
        if mejor_score >= 0.50:
            col_monto = mejor

    if col_fecha is None or col_monto is None:
        return {
            "data": original,
            "fecha": col_fecha,
            "cuenta": col_cuenta,
            "segmento": col_segmento,
            "descripcion": col_descripcion,
            "monto": col_monto,
            "valido": False,
        }

    detalle = original.copy()

    detalle["_fecha"] = pd.to_datetime(
        detalle[col_fecha], errors="coerce", dayfirst=False
    )
    detalle["_mes"] = detalle["_fecha"].dt.to_period("M").dt.to_timestamp()
    detalle["_monto"] = detalle[col_monto].apply(convertir_numero).fillna(0)

    detalle["_cuenta"] = (
        detalle[col_cuenta].astype(str)
        if col_cuenta else ""
    )
    detalle["_segmento"] = (
        detalle[col_segmento].astype(str)
        if col_segmento else ""
    )
    detalle["_descripcion"] = (
        detalle[col_descripcion].astype(str)
        if col_descripcion else ""
    )

    detalle = detalle.dropna(subset=["_mes"]).copy()

    return {
        "data": detalle,
        "fecha": col_fecha,
        "cuenta": col_cuenta,
        "segmento": col_segmento,
        "descripcion": col_descripcion,
        "monto": col_monto,
        "valido": True,
    }


# ============================================================
# ANÁLISIS DE VARIACIONES
# ============================================================
def calcular_variaciones(df_largo):
    if df_largo is None or df_largo.empty:
        return pd.DataFrame()

    base = df_largo.copy()

    base = (
        base.groupby(
            ["pestana", "segmento", "cuenta", "mes"],
            dropna=False,
            as_index=False
        )["monto"]
        .sum()
    )

    base = base.sort_values(
        ["pestana", "segmento", "cuenta", "mes"]
    )

    grupo = ["pestana", "segmento", "cuenta"]

    base["mes_anterior"] = base.groupby(grupo)["monto"].shift(1)
    base["variacion_bs"] = base["monto"] - base["mes_anterior"]

    base["variacion_pct"] = np.where(
        base["mes_anterior"].abs() > 0.01,
        base["variacion_bs"] / base["mes_anterior"].abs() * 100,
        np.nan
    )

    base["promedio_3m"] = (
        base.groupby(grupo)["monto"]
        .transform(lambda s: s.shift(1).rolling(3, min_periods=1).mean())
    )

    base["vs_promedio_3m_bs"] = base["monto"] - base["promedio_3m"]

    base["vs_promedio_3m_pct"] = np.where(
        base["promedio_3m"].abs() > 0.01,
        base["vs_promedio_3m_bs"] / base["promedio_3m"].abs() * 100,
        np.nan
    )

    # Volatilidad histórica y z-score.
    base["media_historica"] = (
        base.groupby(grupo)["monto"]
        .transform(lambda s: s.shift(1).expanding(min_periods=2).mean())
    )
    base["std_historica"] = (
        base.groupby(grupo)["monto"]
        .transform(lambda s: s.shift(1).expanding(min_periods=2).std())
    )

    base["z_score"] = np.where(
        base["std_historica"].abs() > 0.01,
        (base["monto"] - base["media_historica"]) / base["std_historica"],
        np.nan
    )

    return base


def clasificar_variacion(row, umbral_pct, umbral_bs, umbral_z):
    var_bs = abs(row.get("variacion_bs", 0) or 0)
    var_pct = abs(row.get("variacion_pct", np.nan))
    z = abs(row.get("z_score", np.nan))

    criterios = []

    if var_bs >= umbral_bs:
        criterios.append("MONTO")
    if not pd.isna(var_pct) and var_pct >= umbral_pct:
        criterios.append("%")
    if not pd.isna(z) and z >= umbral_z:
        criterios.append("ATIPICA")

    if len(criterios) >= 2:
        return "🔴 CRÍTICA"
    if len(criterios) == 1:
        return "🟠 IMPORTANTE"
    return "🟢 NORMAL"


def obtener_alertas(variaciones, umbral_pct, umbral_bs, umbral_z):
    if variaciones.empty:
        return pd.DataFrame()

    x = variaciones.copy()

    x["nivel"] = x.apply(
        lambda r: clasificar_variacion(
            r, umbral_pct, umbral_bs, umbral_z
        ),
        axis=1
    )

    # Solo interesa el mes más reciente de cada combinación.
    x = x.sort_values("mes")
    idx = (
        x.groupby(["pestana", "segmento", "cuenta"])["mes"]
        .idxmax()
    )
    x = x.loc[idx].copy()

    prioridad = {
        "🔴 CRÍTICA": 0,
        "🟠 IMPORTANTE": 1,
        "🟢 NORMAL": 2,
    }
    x["_prioridad"] = x["nivel"].map(prioridad)
    x = x.sort_values(
        ["_prioridad", "variacion_bs"],
        ascending=[True, False]
    )

    return x.drop(columns=["_prioridad"])


def analizar_detalle(detalle_info):
    if not detalle_info or not detalle_info["valido"]:
        return pd.DataFrame()

    d = detalle_info["data"].copy()

    if d.empty:
        return pd.DataFrame()

    group_cols = ["_mes"]

    if detalle_info["segmento"]:
        group_cols.append("_segmento")
    if detalle_info["cuenta"]:
        group_cols.append("_cuenta")

    resultado = (
        d.groupby(group_cols, dropna=False)["_monto"]
        .agg(["sum", "count"])
        .reset_index()
        .rename(columns={
            "sum": "monto",
            "count": "transacciones",
            "_mes": "mes",
            "_segmento": "segmento",
            "_cuenta": "cuenta",
        })
    )

    if "segmento" not in resultado:
        resultado["segmento"] = "Sin segmento"

    if "cuenta" not in resultado:
        resultado["cuenta"] = "Sin cuenta"

    resultado = resultado.sort_values(["segmento", "cuenta", "mes"])

    grupo = ["segmento", "cuenta"]
    resultado["mes_anterior"] = resultado.groupby(grupo)["monto"].shift(1)
    resultado["variacion_bs"] = (
        resultado["monto"] - resultado["mes_anterior"]
    )
    resultado["variacion_pct"] = np.where(
        resultado["mes_anterior"].abs() > 0.01,
        resultado["variacion_bs"]
        / resultado["mes_anterior"].abs() * 100,
        np.nan
    )

    return resultado


# ============================================================
# EXPORTACIÓN
# ============================================================
def construir_excel_exportable(
    alertas,
    variaciones,
    resumen_segmentos,
    detalle_alertas,
    filtros
):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if not alertas.empty:
            alertas.to_excel(
                writer, sheet_name="Alertas", index=False
            )

        if not variaciones.empty:
            variaciones.to_excel(
                writer, sheet_name="Variaciones", index=False
            )

        if not resumen_segmentos.empty:
            resumen_segmentos.to_excel(
                writer, sheet_name="Segmentos", index=False
            )

        if not detalle_alertas.empty:
            detalle_alertas.to_excel(
                writer, sheet_name="Detalle_Alertas", index=False
            )

        pd.DataFrame([filtros]).to_excel(
            writer, sheet_name="Parametros", index=False
        )

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for col in ws.columns:
                max_len = 0
                letra = col[0].column_letter
                for cell in col[:1000]:
                    try:
                        max_len = max(max_len, len(str(cell.value)))
                    except Exception:
                        pass
                ws.column_dimensions[letra].width = min(max(max_len + 2, 10), 35)

    output.seek(0)
    return output


# ============================================================
# CARGA DEL ARCHIVO
# ============================================================
uploaded = st.file_uploader(
    "📂 Cargar archivo Excel de gastos",
    type=["xlsx", "xls"],
    help="El archivo debe contener las pestañas mensuales/categorizadas y una pestaña Detalle con el histórico de transacciones."
)

if not uploaded:
    st.info(
        "Carga el Excel para comenzar. La aplicación intentará detectar "
        "automáticamente las columnas de meses, cuentas, segmentos y la hoja Detalle."
    )

    with st.expander("¿Qué hará el análisis?"):
        st.markdown("""
        **1. Pestañas de gastos**
        - Identifica columnas mensuales como `2026-09`, `2026-08`, etc.
        - Calcula variación mensual en Bs.
        - Calcula variación porcentual.
        - Compara contra promedio móvil de 3 meses.
        - Detecta comportamientos atípicos mediante z-score.

        **2. Hoja Detalle**
        - Utiliza el histórico de transacciones como respaldo.
        - Permite bajar desde una alerta hasta las operaciones que la originaron.
        - Agrupa por mes, cuenta y segmento.

        **3. Alertas**
        - 🔴 Crítica
        - 🟠 Importante
        - 🟢 Normal
        """)
    st.stop()


# ============================================================
# PROCESAMIENTO
# ============================================================
try:
    xls = leer_excel(uploaded)
except Exception as e:
    st.error(f"No se pudo abrir el archivo: {e}")
    st.stop()

hojas = xls.sheet_names

# Identificación de Detalle.
hoja_detalle = None
for h in hojas:
    if normalizar_texto(h) == "detalle":
        hoja_detalle = h
        break

if hoja_detalle is None:
    candidatos = [
        h for h in hojas
        if "detalle" in normalizar_texto(h)
    ]
    if candidatos:
        hoja_detalle = candidatos[0]

# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.header("⚙️ Parámetros del análisis")

umbral_pct = st.sidebar.number_input(
    "Variación porcentual mínima (%)",
    min_value=1.0,
    max_value=500.0,
    value=30.0,
    step=5.0
)

umbral_bs = st.sidebar.number_input(
    "Variación mínima en Bs.",
    min_value=0.0,
    value=10000.0,
    step=5000.0
)

umbral_z = st.sidebar.number_input(
    "Z-score para comportamiento atípico",
    min_value=1.0,
    max_value=10.0,
    value=2.0,
    step=0.5
)

st.sidebar.divider()
st.sidebar.write("**Hojas detectadas**")
st.sidebar.write(hojas)

# ============================================================
# LEER TODAS LAS PESTAÑAS
# ============================================================
frames = []
diagnostico = []

for hoja in hojas:
    if hoja == hoja_detalle:
        continue

    try:
        df = leer_hoja(xls, hoja)
        largo = convertir_pestana_a_largo(df, hoja)

        if largo is not None and not largo.empty:
            frames.append(largo)
            diagnostico.append({
                "Pestaña": hoja,
                "Filas": len(df),
                "Meses detectados": largo["mes"].nunique(),
                "Estado": "OK"
            })
        else:
            diagnostico.append({
                "Pestaña": hoja,
                "Filas": len(df),
                "Meses detectados": 0,
                "Estado": "No se detectaron meses"
            })
    except Exception as e:
        diagnostico.append({
            "Pestaña": hoja,
            "Filas": 0,
            "Meses detectados": 0,
            "Estado": f"Error: {e}"
        })

if frames:
    todos = pd.concat(frames, ignore_index=True)
else:
    todos = pd.DataFrame()

# Detalle.
detalle_info = None
detalle_df = None

if hoja_detalle:
    try:
        detalle_df = leer_hoja(xls, hoja_detalle)
        detalle_info = preparar_detalle(detalle_df)
    except Exception:
        detalle_info = None

variaciones = calcular_variaciones(todos)

if not variaciones.empty:
    variaciones["nivel"] = variaciones.apply(
        lambda r: clasificar_variacion(
            r, umbral_pct, umbral_bs, umbral_z
        ),
        axis=1
    )

alertas = obtener_alertas(
    variaciones,
    umbral_pct,
    umbral_bs,
    umbral_z
)

detalle_agrupado = analizar_detalle(detalle_info)

# ============================================================
# FILTROS
# ============================================================
st.sidebar.divider()
st.sidebar.header("🔎 Filtros")

if not todos.empty:
    pestanas = sorted(todos["pestana"].dropna().unique())
    seleccion_pestanas = st.sidebar.multiselect(
        "Pestañas",
        pestanas,
        default=pestanas
    )

    segmentos = sorted(
        todos.loc[
            todos["pestana"].isin(seleccion_pestanas),
            "segmento"
        ]
        .dropna()
        .astype(str)
        .unique()
    )

    seleccion_segmentos = st.sidebar.multiselect(
        "Segmentos / cuentas",
        segmentos,
        default=[]
    )

    meses_disponibles = sorted(
        todos["mes"].dropna().unique()
    )

    if meses_disponibles:
        fecha_min = min(meses_disponibles)
        fecha_max = max(meses_disponibles)

        rango = st.sidebar.date_input(
            "Período",
            value=(fecha_min.date(), fecha_max.date())
        )

        if isinstance(rango, tuple) and len(rango) == 2:
            fecha_ini = pd.Timestamp(rango[0])
            fecha_fin = pd.Timestamp(rango[1])
        else:
            fecha_ini = pd.Timestamp(fecha_min)
            fecha_fin = pd.Timestamp(rango)
    else:
        fecha_ini = pd.Timestamp("1900-01-01")
        fecha_fin = pd.Timestamp("2100-01-01")
else:
    seleccion_pestanas = []
    seleccion_segmentos = []
    fecha_ini = pd.Timestamp("1900-01-01")
    fecha_fin = pd.Timestamp("2100-01-01")

# Aplicar filtros.
if not alertas.empty:
    alertas_f = alertas[
        alertas["pestana"].isin(seleccion_pestanas)
        & alertas["mes"].between(fecha_ini, fecha_fin)
    ].copy()

    if seleccion_segmentos:
        alertas_f = alertas_f[
            alertas_f["segmento"].isin(seleccion_segmentos)
        ]
else:
    alertas_f = pd.DataFrame()

if not variaciones.empty:
    var_f = variaciones[
        variaciones["pestana"].isin(seleccion_pestanas)
        & variaciones["mes"].between(fecha_ini, fecha_fin)
    ].copy()

    if seleccion_segmentos:
        var_f = var_f[
            var_f["segmento"].isin(seleccion_segmentos)
        ]
else:
    var_f = pd.DataFrame()

# ============================================================
# KPIs
# ============================================================
st.subheader("📌 Resumen ejecutivo")

c1, c2, c3, c4, c5 = st.columns(5)

total_gasto = todos[
    todos["mes"].between(fecha_ini, fecha_fin)
]["monto"].sum() if not todos.empty else 0

criticas = (
    (alertas_f["nivel"] == "🔴 CRÍTICA").sum()
    if not alertas_f.empty else 0
)

importantes = (
    (alertas_f["nivel"] == "🟠 IMPORTANTE").sum()
    if not alertas_f.empty else 0
)

segmentos_con_alerta = (
    alertas_f["segmento"].nunique()
    if not alertas_f.empty else 0
)

c1.metric("Gasto analizado", f"{total_gasto:,.2f}")
c2.metric("Alertas críticas", f"{criticas:,}")
c3.metric("Alertas importantes", f"{importantes:,}")
c4.metric("Segmentos con alerta", f"{segmentos_con_alerta:,}")
c5.metric("Pestañas analizadas", f"{len(seleccion_pestanas):,}")

# ============================================================
# TABS PRINCIPALES
# ============================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🚨 Alertas",
    "📈 Variación mensual",
    "🧩 Segmentos",
    "🔍 Detalle de transacciones",
    "🛠 Diagnóstico"
])

# ============================================================
# TAB 1 - ALERTAS
# ============================================================
with tab1:
    st.subheader("🚨 Principales variaciones detectadas")

    if alertas_f.empty:
        st.success("No hay variaciones que superen los parámetros establecidos.")
    else:
        mostrar = alertas_f.copy()

        columnas = [
            "nivel", "pestana", "segmento", "cuenta", "mes",
            "monto", "mes_anterior", "variacion_bs",
            "variacion_pct", "promedio_3m",
            "vs_promedio_3m_pct", "z_score"
        ]

        columnas = [c for c in columnas if c in mostrar.columns]

        st.dataframe(
            mostrar[columnas].style.format({
                c: "{:,.2f}" for c in [
                    "monto", "mes_anterior", "variacion_bs",
                    "variacion_pct", "promedio_3m",
                    "vs_promedio_3m_pct", "z_score"
                ] if c in mostrar.columns
            }),
            use_container_width=True,
            height=500
        )

        st.markdown("### ¿Qué significa cada alerta?")
        st.markdown("""
        - **🔴 CRÍTICA:** cumple al menos dos señales de riesgo.
        - **🟠 IMPORTANTE:** cumple al menos una señal.
        - **🟢 NORMAL:** no supera los parámetros definidos.

        La evaluación combina:
        1. Variación absoluta en Bs.
        2. Variación porcentual respecto al mes anterior.
        3. Comparación contra promedio móvil de 3 meses.
        4. Comportamiento estadísticamente atípico mediante z-score.
        """)

# ============================================================
# TAB 2 - VARIACIÓN MENSUAL
# ============================================================
with tab2:
    st.subheader("📈 Evolución mensual")

    if var_f.empty:
        st.info("No hay información para los filtros seleccionados.")
    else:
        opciones = (
            var_f["pestana"] + " | "
            + var_f["segmento"].astype(str)
            + " | "
            + var_f["cuenta"].astype(str)
        ).unique()

        seleccion = st.selectbox(
            "Seleccione cuenta / segmento para ver su evolución",
            opciones
        )

        partes = seleccion.split(" | ", 2)
        p, s, c = partes

        serie = var_f[
            (var_f["pestana"] == p)
            & (var_f["segmento"].astype(str) == s)
            & (var_f["cuenta"].astype(str) == c)
        ].sort_values("mes")

        if not serie.empty:
            graf = serie.set_index("mes")[["monto", "promedio_3m"]]
            st.line_chart(graf)

            st.dataframe(
                serie[
                    [
                        "mes", "monto", "mes_anterior",
                        "variacion_bs", "variacion_pct",
                        "promedio_3m", "vs_promedio_3m_pct",
                        "z_score", "nivel"
                    ]
                ].style.format({
                    c: "{:,.2f}" for c in [
                        "monto", "mes_anterior", "variacion_bs",
                        "variacion_pct", "promedio_3m",
                        "vs_promedio_3m_pct", "z_score"
                    ]
                }),
                use_container_width=True
            )

# ============================================================
# TAB 3 - SEGMENTOS
# ============================================================
with tab3:
    st.subheader("🧩 Análisis por segmento")

    if var_f.empty:
        st.info("No hay datos para los filtros seleccionados.")
    else:
        resumen_segmentos = (
            var_f.groupby(
                ["pestana", "segmento", "mes"],
                dropna=False,
                as_index=False
            )
            .agg(
                gasto=("monto", "sum"),
                variacion=("variacion_bs", "sum"),
            )
        )

        resumen_segmentos["variacion_pct"] = np.where(
            (resumen_segmentos["gasto"] - resumen_segmentos["variacion"]).abs() > 0.01,
            resumen_segmentos["variacion"]
            / (resumen_segmentos["gasto"] - resumen_segmentos["variacion"]).abs()
            * 100,
            np.nan
        )

        ultimo_mes = resumen_segmentos["mes"].max()

        resumen_ultimo = resumen_segmentos[
            resumen_segmentos["mes"] == ultimo_mes
        ].copy()

        resumen_ultimo = resumen_ultimo.sort_values(
            "variacion",
            key=lambda s: s.abs(),
            ascending=False
        )

        st.markdown(
            f"**Último mes disponible en el filtro:** "
            f"{ultimo_mes.strftime('%Y-%m') if not pd.isna(ultimo_mes) else '-'}"
        )

        st.dataframe(
            resumen_ultimo.style.format({
                "gasto": "{:,.2f}",
                "variacion": "{:,.2f}",
                "variacion_pct": "{:,.2f}%"
            }),
            use_container_width=True,
            height=450
        )

        seg = st.selectbox(
            "Seleccione un segmento para visualizar su tendencia",
            sorted(resumen_segmentos["segmento"].astype(str).unique())
        )

        serie_seg = resumen_segmentos[
            resumen_segmentos["segmento"].astype(str) == seg
        ]

        graf_seg = (
            serie_seg.groupby("mes")["gasto"]
            .sum()
            .sort_index()
        )

        st.line_chart(graf_seg)

# ============================================================
# TAB 4 - DETALLE
# ============================================================
with tab4:
    st.subheader("🔍 Auditoría de transacciones")

    if detalle_info is None:
        st.warning(
            "No se encontró una hoja Detalle. "
            "La aplicación puede analizar las pestañas resumen, "
            "pero no podrá bajar hasta la transacción."
        )

    elif not detalle_info["valido"]:
        st.warning(
            "La hoja Detalle fue encontrada, pero no se pudieron identificar "
            "automáticamente una columna de Fecha y una columna de Monto."
        )

        st.write("Columnas detectadas:")
        st.write({
            "Fecha": detalle_info["fecha"],
            "Cuenta": detalle_info["cuenta"],
            "Segmento": detalle_info["segmento"],
            "Descripción": detalle_info["descripcion"],
            "Monto": detalle_info["monto"],
        })

    else:
        d = detalle_info["data"].copy()

        col1, col2, col3 = st.columns(3)

        meses_det = sorted(d["_mes"].dropna().unique())
        segmentos_det = sorted(
            d["_segmento"].astype(str).unique()
        )

        with col1:
            mes_det = st.selectbox(
                "Mes",
                ["Todos"] + [
                    pd.Timestamp(x).strftime("%Y-%m")
                    for x in meses_det
                ]
            )

        with col2:
            segmento_det = st.selectbox(
                "Segmento",
                ["Todos"] + segmentos_det
            )

        with col3:
            texto = st.text_input(
                "Buscar en descripción / cuenta"
            )

        filtrado = d.copy()

        if mes_det != "Todos":
            filtrado = filtrado[
                filtrado["_mes"].dt.strftime("%Y-%m") == mes_det
            ]

        if segmento_det != "Todos":
            filtrado = filtrado[
                filtrado["_segmento"].astype(str) == segmento_det
            ]

        if texto:
            mask = (
                filtrado["_descripcion"].astype(str)
                .str.contains(texto, case=False, na=False)
                |
                filtrado["_cuenta"].astype(str)
                .str.contains(texto, case=False, na=False)
            )
            filtrado = filtrado[mask]

        st.metric(
            "Transacciones encontradas",
            f"{len(filtrado):,}"
        )

        columnas_mostrar = []

        if detalle_info["fecha"]:
            columnas_mostrar.append(detalle_info["fecha"])
        if detalle_info["cuenta"]:
            columnas_mostrar.append(detalle_info["cuenta"])
        if detalle_info["segmento"]:
            columnas_mostrar.append(detalle_info["segmento"])
        if detalle_info["descripcion"]:
            columnas_mostrar.append(detalle_info["descripcion"])
        if detalle_info["monto"]:
            columnas_mostrar.append(detalle_info["monto"])

        # Agrega columnas internas útiles.
        mostrar_detalle = filtrado[columnas_mostrar].copy()

        st.dataframe(
            mostrar_detalle,
            use_container_width=True,
            height=500
        )

        if not alertas_f.empty:
            st.markdown("### Transacciones relacionadas con una alerta")

            alerta_idx = st.selectbox(
                "Seleccione una alerta",
                range(len(alertas_f)),
                format_func=lambda i: (
                    f"{alertas_f.iloc[i]['pestana']} | "
                    f"{alertas_f.iloc[i]['segmento']} | "
                    f"{alertas_f.iloc[i]['cuenta']} | "
                    f"{pd.Timestamp(alertas_f.iloc[i]['mes']).strftime('%Y-%m')} | "
                    f"{alertas_f.iloc[i]['nivel']}"
                )
            )

            alerta = alertas_f.iloc[alerta_idx]

            rel = d.copy()

            rel = rel[
                rel["_mes"] == alerta["mes"]
            ]

            if detalle_info["segmento"]:
                rel = rel[
                    rel["_segmento"].astype(str)
                    == str(alerta["segmento"])
                ]

            if detalle_info["cuenta"]:
                rel = rel[
                    rel["_cuenta"].astype(str)
                    == str(alerta["cuenta"])
                ]

            st.write(
                f"**Monto de la alerta:** {alerta['monto']:,.2f}  |  "
                f"**Variación:** {alerta['variacion_bs']:,.2f}  |  "
                f"**Variación %:** "
                f"{alerta['variacion_pct']:,.2f}%"
                if not pd.isna(alerta["variacion_pct"])
                else
                f"**Monto de la alerta:** {alerta['monto']:,.2f}  |  "
                f"**Variación:** {alerta['variacion_bs']:,.2f}"
            )

            if rel.empty:
                st.info(
                    "No se encontraron transacciones con la misma combinación "
                    "de mes + segmento + cuenta."
                )
            else:
                cols_rel = []
                for c in [
                    detalle_info["fecha"],
                    detalle_info["cuenta"],
                    detalle_info["segmento"],
                    detalle_info["descripcion"],
                    detalle_info["monto"],
                ]:
                    if c and c not in cols_rel:
                        cols_rel.append(c)

                st.dataframe(
                    rel[cols_rel],
                    use_container_width=True,
                    height=400
                )

# ============================================================
# TAB 5 - DIAGNÓSTICO
# ============================================================
with tab5:
    st.subheader("🛠 Diagnóstico del archivo")

    st.write(
        "Esta tabla permite comprobar qué pestañas fueron utilizadas "
        "y qué estructura detectó automáticamente el programa."
    )

    st.dataframe(
        pd.DataFrame(diagnostico),
        use_container_width=True
    )

    st.markdown("### Hoja Detalle")

    st.write({
        "Hoja encontrada": hoja_detalle,
        "Fecha": detalle_info["fecha"] if detalle_info else None,
        "Cuenta": detalle_info["cuenta"] if detalle_info else None,
        "Segmento": detalle_info["segmento"] if detalle_info else None,
        "Descripción": detalle_info["descripcion"] if detalle_info else None,
        "Monto": detalle_info["monto"] if detalle_info else None,
    })

    st.markdown("### Pestañas analizadas")
    if not todos.empty:
        st.write(sorted(todos["pestana"].unique()))
        st.write(
            f"Total de registros mensuales procesados: "
            f"{len(todos):,}"
        )

# ============================================================
# EXPORTACIÓN
# ============================================================
st.divider()
st.subheader("📥 Exportar análisis")

if "resumen_segmentos" not in locals():
    resumen_segmentos = pd.DataFrame()

filtros_export = {
    "umbral_pct": umbral_pct,
    "umbral_bs": umbral_bs,
    "umbral_z": umbral_z,
    "fecha_inicio": str(fecha_ini.date()),
    "fecha_fin": str(fecha_fin.date()),
    "pestanas": ", ".join(seleccion_pestanas),
    "segmentos": ", ".join(seleccion_segmentos),
}

excel = construir_excel_exportable(
    alertas_f,
    var_f,
    resumen_segmentos,
    detalle_agrupado,
    filtros_export
)

st.download_button(
    "📥 Descargar reporte Excel",
    data=excel,
    file_name=f"Analisis_Variaciones_Gastos_{datetime.now():%Y%m%d_%H%M}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
