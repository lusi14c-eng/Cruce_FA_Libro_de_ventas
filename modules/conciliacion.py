import io
import pandas as pd
import numpy as np
import streamlit as st
import re
from utils.excel_exporter import generar_excel_resultado

TIPOS_VALIDOS = ["FAC", "N/C", "N/D"]

def normalizar_documento(valor):
    if pd.isna(valor):
        return ""
    s = str(valor).strip().upper()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    s = re.sub(r"[\s\-_./]+", "", s)
    if s.isdigit():
        s = s.lstrip("0") or "0"
    return s

def convertir_numero(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float, np.integer, np.floating)):
        return float(valor)
    s = str(valor).strip().replace(" ", "")
    s = re.sub(r"[^\d,\.\-\(\)]", "", s)
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        partes = s.split(",")
        if len(partes[-1]) in (1, 2):
            s = "".join(partes[:-1]) + "." + partes[-1]
        else:
            s = s.replace(",", "")
    elif s.count(".") > 1:
        partes = s.split(".")
        s = "".join(partes[:-1]) + "." + partes[-1]
    try:
        return float(s)
    except Exception:
        return 0.0

@st.cache_data(show_spinner=False)
def cargar_listado_facturacion(file_bytes, file_name, hoja="Documentos_CC"):
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    hoja_uso = hoja if hoja in xls.sheet_names else xls.sheet_names[0]
    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=hoja_uso, engine="openpyxl")
    df = pd.DataFrame()

    df["doc_num"] = df_raw.iloc[:, 0].apply(normalizar_documento)
    df["tipo"] = df_raw.iloc[:, 1].astype(str).str.strip().str.upper().replace({
        "FACTURA": "FAC", "FACT": "FAC", "F": "FAC",
        "NC": "N/C", "NOTACREDITO": "N/C",
        "ND": "N/D", "NOTADEBITO": "N/D"
    })
    df["clave"] = df["tipo"] + "|" + df["doc_num"]
    df["fecha_lista"] = pd.to_datetime(df_raw.iloc[:, 2], errors="coerce")
    df["codigo_cliente"] = df_raw.iloc[:, 5].fillna("").astype(str)
    df["cliente"] = df_raw.iloc[:, 6].fillna("").astype(str)

    # Base Imponible tomada de Columna V (Subtotal-Descuento Bs.)
    base = df_raw.iloc[:, 21].apply(convertir_numero)
    iva = df_raw.iloc[:, 18].apply(convertir_numero)
    exento = df_raw.iloc[:, 19].apply(convertir_numero)

    mask_nc = df["tipo"] == "N/C"
    base.loc[mask_nc] = base.loc[mask_nc].abs() * -1
    iva.loc[mask_nc] = iva.loc[mask_nc].abs() * -1
    exento.loc[mask_nc] = exento.loc[mask_nc].abs() * -1

    df["base_lista"] = base
    df["iva_lista"] = iva
    df["exento_lista"] = exento
    df["total_lista"] = base + iva + exento

    return df[(df["doc_num"] != "") & (df["tipo"].isin(TIPOS_VALIDOS))].copy()

@st.cache_data(show_spinner=False)
def cargar_libro_ventas(file_bytes, file_name):
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    sheet_name = xls.sheet_names[0]
    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, skiprows=5, engine="openpyxl")

    col_fac = df_raw.iloc[:, 5].apply(normalizar_documento)
    col_nd = df_raw.iloc[:, 8].apply(normalizar_documento)
    col_nc = df_raw.iloc[:, 9].apply(normalizar_documento)

    doc_num = col_fac.replace("", np.nan).combine_first(col_nd.replace("", np.nan)).combine_first(col_nc.replace("", np.nan)).fillna("")
    tipo = np.where(col_nc != "", "N/C", np.where(col_nd != "", "N/D", np.where(col_fac != "", "FAC", "")))

    df = pd.DataFrame()
    df["doc_num"] = doc_num
    df["tipo"] = tipo
    df["clave"] = df["tipo"] + "|" + df["doc_num"]
    df["fecha_libro"] = pd.to_datetime(df_raw.iloc[:, 1], errors="coerce")
    df["rif_libro"] = df_raw.iloc[:, 2].fillna("").astype(str)
    df["cliente_libro"] = df_raw.iloc[:, 3].fillna("").astype(str)

    col_num_comprobante = df_raw.iloc[:, 13].fillna("").astype(str).str.strip()
    col_iva_retenido = df_raw.iloc[:, 26].apply(convertir_numero)

    exento = df_raw.iloc[:, 15].apply(convertir_numero)
    base = df_raw.iloc[:, 17].apply(convertir_numero)
    iva = df_raw.iloc[:, 19].apply(convertir_numero)
    total = df_raw.iloc[:, 14].apply(convertir_numero)

    es_retencion = (base == 0) & ((col_num_comprobante != "") | (col_iva_retenido > 0))
    df["es_retencion"] = es_retencion
    df["num_comprobante"] = col_num_comprobante
    df["iva_retenido"] = col_iva_retenido

    total = np.where(total == 0, base + iva + exento, total)

    mask_nc = df["tipo"] == "N/C"
    base = np.where(mask_nc, -np.abs(base), base)
    iva = np.where(mask_nc, -np.abs(iva), iva)
    exento = np.where(mask_nc, -np.abs(exento), exento)
    total = np.where(mask_nc, -np.abs(total), total)

    df["base_libro"] = base
    df["iva_libro"] = iva
    df["exento_libro"] = exento
    df["total_libro"] = total
    df["periodo"] = file_name.replace(".xlsx", "").replace(".xlsm", "")

    return df[(df["doc_num"] != "") & (df["tipo"].isin(TIPOS_VALIDOS))].copy()

def conciliar_datos(df_lista, df_libros, tolerancia=0.50):
    df_libros_ventas = df_libros[~df_libros["es_retencion"]].copy()
    df_libros_retenciones = df_libros[df_libros["es_retencion"]].copy()

    df_lista["dup_lista"] = df_lista.groupby("clave")["clave"].transform("count") > 1
    df_libros_ventas["dup_libro"] = df_libros_ventas.groupby("clave")["clave"].transform("count") > 1

    lista_agg = df_lista.groupby("clave", as_index=False).agg(
        tipo=("tipo", "first"), doc_num=("doc_num", "first"), fecha_lista=("fecha_lista", "min"),
        codigo_cliente=("codigo_cliente", "first"), cliente=("cliente", "first"),
        base_lista=("base_lista", "sum"), iva_lista=("iva_lista", "sum"),
        exento_lista=("exento_lista", "sum"), total_lista=("total_lista", "sum"),
        dup_lista=("dup_lista", "max"), ocurrencias_lista=("clave", "size")
    )

    libro_agg = df_libros_ventas.groupby("clave", as_index=False).agg(
        tipo=("tipo", "first"), doc_num=("doc_num", "first"), fecha_libro=("fecha_libro", "min"),
        cliente_libro=("cliente_libro", "first"), base_libro=("base_libro", "sum"),
        iva_libro=("iva_libro", "sum"), exento_libro=("exento_libro", "sum"),
        total_libro=("total_libro", "sum"), dup_libro=("dup_libro", "max"),
        periodo=("periodo", lambda x: " + ".join(sorted(set(str(v) for v in x if pd.notna(v))))),
        ocurrencias_libro=("clave", "size")
    )

    ret_agg = df_libros_retenciones.groupby("clave", as_index=False).agg(
        num_comprobante=("num_comprobante", lambda x: " + ".join(sorted(set(str(v) for v in x if pd.notna(v) and str(v) != "")))),
        iva_retenido=("iva_retenido", "sum"), tiene_retencion=("es_retencion", "any")
    )

    res = pd.merge(lista_agg, libro_agg, on="clave", how="outer", suffixes=("_lis", "_lib"))
    res = pd.merge(res, ret_agg, on="clave", how="left")

    res["tipo"] = res["tipo_lis"].fillna(res["tipo_lib"])
    res["documento"] = res["doc_num_lis"].fillna(res["doc_num_lib"])
    res["periodo"] = res["periodo"].fillna("No en libro")
    res["tiene_retencion"] = res["tiene_retencion"].fillna(False)

    for col in ["base_lista", "iva_lista", "exento_lista", "total_lista", "base_libro", "iva_libro", "exento_libro", "total_libro", "iva_retenido"]:
        res[col] = res[col].fillna(0.0)

    res["diferencia_exento"] = (res["exento_libro"] - res["exento_lista"]).round(2)
    res["diferencia_base"] = (res["base_libro"] - res["base_lista"]).round(2)
    res["diferencia_iva"] = (res["iva_libro"] - res["iva_lista"]).round(2)
    res["diferencia_total"] = (res["total_libro"] - res["total_lista"]).round(2)

    existe_lista = res["ocurrencias_lista"].notna() & (res["ocurrencias_lista"] > 0)
    existe_libro = res["ocurrencias_libro"].notna() & (res["ocurrencias_libro"] > 0)

    res["estado"] = "OK"
    res.loc[existe_lista & ~existe_libro, "estado"] = "FALTA EN LIBRO"
    res.loc[~existe_lista & existe_libro, "estado"] = "SOLO EN LIBRO"

    ambos = existe_lista & existe_libro
    res.loc[ambos & (res["diferencia_total"].abs() > tolerancia), "estado"] = "DIFERENCIA DE MONTO"
    res.loc[ambos & (res["diferencia_total"].abs() <= tolerancia), "estado"] = "OK"

    es_dup = res["dup_lista"].fillna(False) | res["dup_libro"].fillna(False)
    res.loc[es_dup, "estado"] = res.loc[es_dup, "estado"] + " / DUPLICADO"

    return res

def modulo_conciliacion(sucursal):
    st.title(f"📊 Conciliación de Ventas - {sucursal}")

    f_lista = st.file_uploader("1. Cargar Lista de Facturación (Excel)", type=["xlsx", "xlsm"], key=f"lista_{sucursal}")
    f_libros = st.file_uploader("2. Cargar Libro(s) de Ventas (Excel)", type=["xlsx", "xlsm"], accept_multiple_files=True, key=f"libros_{sucursal}")

    tolerancia = st.sidebar.number_input("Tolerancia monetaria (Bs.)", min_value=0.0, value=0.50, step=0.01)

    if not f_lista or not f_libros:
        st.info("👋 Sube los archivos requeridos para iniciar la conciliación.")
        return

    if st.button("🚀 Ejecutar Conciliación", type="primary", use_container_width=True):
        with st.spinner("⏳ Procesando conciliación..."):
            df_lista = cargar_listado_facturacion(f_lista.getvalue(), f_lista.name)
            dict_libros = {f.name: cargar_libro_ventas(f.getvalue(), f.name) for f in f_libros}
            df_libros_todos = pd.concat(dict_libros.values(), ignore_index=True)

            res = conciliar_datos(df_lista, df_libros_todos, tolerancia)

        st.success("✅ Conciliación realizada.")
        
        # Muestra métricas
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total", len(res))
        m2.metric("OK", int(res["estado"].str.startswith("OK").sum()))
        m3.metric("Faltan en Libro", int(res["estado"].str.startswith("FALTA EN LIBRO").sum()))
        m4.metric("Solo en Libro", int(res["estado"].str.startswith("SOLO EN LIBRO").sum()))
        m5.metric("Diferencias", int(res["estado"].str.startswith("DIFERENCIA DE MONTO").sum()))

        st.dataframe(res, use_container_width=True)

        excel_bytes = generar_excel_resultado(res, df_lista, dict_libros, tolerancia)
        st.download_button(
            "📥 Descargar Reporte en Excel",
            data=excel_bytes,
            file_name=f"Conciliacion_{sucursal}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
