import csv
import io
import re
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

# ============================================================
# CUENTAS CONTABLES DEFAULT
# ============================================================
CUENTA_INGRESOS = "4.1.1"
CUENTA_EXENTO = "7.1.3.45.1.997"
CUENTA_BASE_GRAVADA = "2.1.3.04.1.001"
CUENTA_RETENCION_IVA = "2.1.3.04.1.006"


# ============================================================
# FUNCIONES GENERALES Y LIMPIEZA
# ============================================================

def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor).strip()
    valor = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode("ascii")
    return valor.upper().strip()


def limpiar_numero(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float, np.number)):
        return float(valor)
    texto = str(valor).strip()
    if texto == "":
        return 0.0

    texto = texto.replace("$", "").replace("Bs.", "").replace("Bs", "").replace(" ", "")

    if "," in texto and "." in texto:
        ultima_coma = texto.rfind(",")
        ultimo_punto = texto.rfind(".")
        if ultima_coma > ultimo_punto:
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(",", ".")

    try:
        return float(texto)
    except Exception:
        return 0.0


def limpiar_rif(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    return texto.replace(".", "").replace(" ", "").replace("-", "")


def limpiar_documento(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto.zfill(7) if texto.isdigit() else texto


def tipo_documento(valor):
    texto = normalizar_texto(valor)
    if texto in ["FAC", "FACTURA", "FACT"]:
        return "FAC"
    if texto in ["N/C", "NC", "NOTA DE CREDITO", "NOTA CREDITO"]:
        return "N/C"
    if texto in ["N/D", "ND", "NOTA DE DEBITO", "NOTA DEBITO"]:
        return "N/D"
    return texto


def separar_fuente(valor):
    if pd.isna(valor):
        return "", ""
    texto = str(valor).strip().upper()
    patrones = [
        r"^(FAC)#?(.*)$",
        r"^(N/C)#?(.*)$",
        r"^(N/D)#?(.*)$",
        r"^(NC)#?(.*)$",
        r"^(ND)#?(.*)$",
    ]
    for patron in patrones:
        match = re.match(patron, texto)
        if match:
            tipo = match.group(1)
            numero = match.group(2).strip()
            tipo = tipo_documento(tipo)
            return tipo, numero
    return "", texto


def detectar_columna(df, candidatos):
    columnas_normalizadas = {normalizar_texto(col): col for col in df.columns}
    for candidato in candidatos:
        candidato_norm = normalizar_texto(candidato)
        if candidato_norm in columnas_normalizadas:
            return columnas_normalizadas[candidato_norm]

    for columna_norm, columna_real in columnas_normalizadas.items():
        for candidato in candidatos:
            candidato_norm = normalizar_texto(candidato)
            if candidato_norm in columna_norm:
                return columna_real
    return None


def leer_auxiliar_fa(uploaded_file):
    nombre = uploaded_file.name.lower()
    contenido = uploaded_file.read()

    if nombre.endswith((".txt", ".csv", ".dat")):
        texto = contenido.decode("utf-8", errors="replace")
        try:
            df = pd.read_csv(
                io.StringIO(texto),
                sep="|",
                dtype=str,
                engine="python",
                quoting=csv.QUOTE_NONE,
                on_bad_lines="skip"
            )
        except Exception:
            df = pd.read_csv(
                io.StringIO(texto),
                sep="|",
                dtype=str,
                engine="c",
                on_bad_lines="skip"
            )
    elif nombre.endswith((".xlsx", ".xls")):
        df = pd.read_excel(io.BytesIO(contenido), dtype=str)
    else:
        raise ValueError("Formato de Auxiliar FA no soportado.")

    df = df.dropna(axis=1, how="all")
    df.columns = [str(col).strip().strip('"') for col in df.columns]
    df = df.apply(lambda col: col.str.strip('"') if col.dtype == "object" else col)
    return df


def preparar_auxiliar_fa(df):
    df = df.copy()

    mapa = {
        "numero": detectar_columna(df, [
            "Numero", "Número", "Nro", "Nro.", "Num", "Num.", "Documento", "DOCUMENTO", 
            "Nro Documento", "Nro. Documento", "Nro Doc", "Nro. Doc", "FACTURA", "DOC_NUM"
        ]),
        "tipo": detectar_columna(df, ["Tipo", "Tipo Doc", "Tipo Documento", "TIP_DOC", "TIPO_DOC"]),
        "fecha": detectar_columna(df, ["Fecha Doc", "Fecha", "Fecha Documento", "FEC_DOC", "FECHA_DOC"]),
        "anulada": detectar_columna(df, ["ANULADA", "Anulada", "Anulado", "Status", "Estado", "ANULADO"]),
        "moneda": detectar_columna(df, ["moneda_factura", "Moneda", "MONEDA"]),
        "subtipo": detectar_columna(df, ["Sub Tipo", "Subtipo", "SUB_TIPO"]),
        "factura_afectada": detectar_columna(df, ["Factura Afectada", "Doc Afectado", "Afectado", "FACTURA_AFECTADA","U_NUM_FAC_AFECTADA"]),
        "cliente": detectar_columna(df, ["Cliente", "Cod Cliente", "Código Cliente", "COD_CLI"]),
        "nombre": detectar_columna(df, ["Nombre", "Razon Social", "Razón Social", "Cliente", "NOM_CLI", "NOMBRE_CLIENTE]),
        "asiento": detectar_columna(df, ["Asiento", "Nro Asiento"]),
        "modulo": detectar_columna(df, ["Modulo", "Módulo"]),
        "tipo_cambio": detectar_columna(df, ["Tipo Cambio", "Tasa", "Tasa Cambio"]),
        "costo_total_bs": detectar_columna(df, ["Costo Total Bs.", "Costo Total Bs", "Costo Total"]),
        "base_bs": detectar_columna(df, [
            "Base Imponible Bs.", "Base Imponible Bs", "Base Imponible", "Base Bs.", "Base Bs",
            "Base Gravada", "Monto Base", "BASE_IMP", "M_BASE", "BASE_IMP_IVA_Bs.", "BASE_IMP_IVA_Bs"
        ]),
        "iva_bs": detectar_columna(df, ["Iva Bs.", "IVA Bs.", "Iva Bs", "IVA Bs", "Monto IVA", "IVA"]),
        "flete_bs": detectar_columna(df, ["Flete Bs.", "Flete Bs", "Flete"]),
        "monto_bs": detectar_columna(df, ["Monto Bs.", "Monto Bs", "Total Bs.", "Total Bs", "Total", "Monto Total"]),
        "base_usd": detectar_columna(df, ["Base Imponible $", "Base Imponible USD", "Base $"]),
        "iva_usd": detectar_columna(df, ["Iva $", "IVA $", "IVA USD"]),
        "monto_usd": detectar_columna(df, ["Monto $", "Monto USD", "Total $"]),
    }

    obligatorias = ["numero", "tipo", "fecha", "nombre", "base_bs", "iva_bs", "monto_bs"]
    faltantes = [campo for campo in obligatorias if mapa[campo] is None]

    if faltantes:
        raise ValueError(f"No se encontraron las columnas necesarias del Auxiliar FA: {faltantes}")

    salida = pd.DataFrame()
    salida["Numero"] = df[mapa["numero"]].apply(limpiar_documento)
    salida["Tipo"] = df[mapa["tipo"]].apply(tipo_documento)
    salida["Fecha"] = pd.to_datetime(df[mapa["fecha"]], errors="coerce")
    salida["Anulada"] = df[mapa["anulada"]].astype(str).str.upper().str.strip() if mapa["anulada"] else ""
    salida["Moneda"] = df[mapa["moneda"]].astype(str) if mapa["moneda"] else ""
    salida["Sub Tipo"] = df[mapa["subtipo"]].astype(str) if mapa["subtipo"] else ""
    salida["Factura Afectada"] = df[mapa["factura_afectada"]].apply(limpiar_documento) if mapa["factura_afectada"] else ""
    salida["Cliente"] = df[mapa["cliente"]].astype(str) if mapa["cliente"] else ""
    salida["Nombre"] = df[mapa["nombre"]].astype(str)
    salida["Asiento"] = df[mapa["asiento"]].astype(str) if mapa["asiento"] else ""
    salida["Modulo"] = df[mapa["modulo"]].astype(str) if mapa["modulo"] else ""
    salida["Tipo Cambio"] = df[mapa["tipo_cambio"]].apply(limpiar_numero) if mapa["tipo_cambio"] else 0
    salida["Costo Total Bs"] = df[mapa["costo_total_bs"]].apply(limpiar_numero) if mapa["costo_total_bs"] else 0
    salida["Base Imponible Bs"] = df[mapa["base_bs"]].apply(limpiar_numero)
    salida["IVA Bs"] = df[mapa["iva_bs"]].apply(limpiar_numero)
    salida["Flete Bs"] = df[mapa["flete_bs"]].apply(limpiar_numero) if mapa["flete_bs"] else 0
    salida["Monto Bs"] = df[mapa["monto_bs"]].apply(limpiar_numero)
    salida["Base Imponible $"] = df[mapa["base_usd"]].apply(limpiar_numero) if mapa["base_usd"] else 0
    salida["IVA $"] = df[mapa["iva_usd"]].apply(limpiar_numero) if mapa["iva_usd"] else 0
    salida["Monto $"] = df[mapa["monto_usd"]].apply(limpiar_numero) if mapa["monto_usd"] else 0

    salida["Clave Documento"] = salida["Tipo"].astype(str) + "|" + salida["Numero"].astype(str)
    salida["Alicuota"] = np.where(salida["Base Imponible Bs"] != 0, (salida["IVA Bs"] / salida["Base Imponible Bs"] * 100), 0)
    salida["Alicuota"] = salida["Alicuota"].round(2)
    salida["Estado"] = np.where(salida["Anulada"].isin(["S", "SI", "Y", "YES"]), "ANULADA", "ACTIVA")

    return salida


def leer_mayor(uploaded_file):
    nombre = uploaded_file.name.lower()
    contenido = uploaded_file.read()

    if nombre.endswith((".xlsx", ".xls")):
        df = pd.read_excel(io.BytesIO(contenido), dtype=str)
    elif nombre.endswith(".csv"):
        try:
            df = pd.read_csv(io.BytesIO(contenido), dtype=str)
        except Exception:
            df = pd.read_csv(io.BytesIO(contenido), sep=";", dtype=str)
    elif nombre.endswith(".txt"):
        try:
            df = pd.read_csv(io.BytesIO(contenido), sep="|", dtype=str, engine="python")
        except Exception:
            df = pd.read_csv(io.BytesIO(contenido), sep="\t", dtype=str, engine="python")
    else:
        raise ValueError("Formato de mayor no soportado.")

    df = df.dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def preparar_mayor(df):
    df = df.copy()

    col_asiento = detectar_columna(df, ["Asiento"])
    col_fecha = detectar_columna(df, ["Fecha"])
    col_cuenta = detectar_columna(df, ["Cuenta Contable"])
    col_descripcion = detectar_columna(df, ["Descripción de la cuenta", "Descripcion de la cuenta"])
    col_fuente = detectar_columna(df, ["Fuente"])
    col_referencia = detectar_columna(df, ["Referencia"])
    col_debito = detectar_columna(df, ["Débito VES", "Debito VES"])
    col_credito = detectar_columna(df, ["Crédito VES", "Credito VES"])
    col_nit = detectar_columna(df, ["Nit", "NIT", "RIF"])

    obligatorias = {
        "Asiento": col_asiento,
        "Fecha": col_fecha,
        "Cuenta Contable": col_cuenta,
        "Fuente": col_fuente,
        "Débito VES": col_debito,
        "Crédito VES": col_credito,
        "NIT": col_nit,
    }

    faltantes = [nombre for nombre, columna in obligatorias.items() if columna is None]
    if faltantes:
        raise ValueError("El mayor no contiene las columnas requeridas: " + ", ".join(faltantes))

    salida = pd.DataFrame()
    salida["Asiento"] = df[col_asiento].astype(str)
    salida["Fecha"] = pd.to_datetime(df[col_fecha], errors="coerce")
    salida["Cuenta Contable"] = df[col_cuenta].astype(str).str.strip()
    salida["Descripción de la cuenta"] = df[col_descripcion].astype(str) if col_descripcion else ""
    salida["Fuente"] = df[col_fuente].astype(str).str.strip()
    salida["Referencia"] = df[col_referencia].astype(str).str.strip() if col_referencia else ""
    salida["Débito VES"] = df[col_debito].apply(limpiar_numero)
    salida["Crédito VES"] = df[col_credito].apply(limpiar_numero)
    salida["NIT"] = df[col_nit].apply(limpiar_rif)
    salida["Movimiento VES"] = salida["Crédito VES"] - salida["Débito VES"]

    parsed = salida["Fuente"].apply(separar_fuente)
    salida["Tipo Documento"] = parsed.apply(lambda x: x[0])
    salida["Numero Documento"] = parsed.apply(lambda x: limpiar_documento(x[1]))
    salida["Clave Documento"] = salida["Tipo Documento"] + "|" + salida["Numero Documento"]

    return salida


def resumir_mayor(mayor):
    ingresos = mayor[mayor["Cuenta Contable"].str.startswith(CUENTA_INGRESOS)].copy()
    ingresos = ingresos.groupby(["Clave Documento", "Tipo Documento", "Numero Documento", "NIT"], dropna=False)["Movimiento VES"].sum().reset_index().rename(columns={"Movimiento VES": "Ingreso Contabilidad"})

    base = mayor[mayor["Cuenta Contable"].astype(str) == CUENTA_BASE_GRAVADA].copy()
    base = base.groupby(["Clave Documento", "Tipo Documento", "Numero Documento", "NIT"], dropna=False)["Movimiento VES"].sum().reset_index().rename(columns={"Movimiento VES": "Base Contabilidad"})

    exento = mayor[mayor["Cuenta Contable"].astype(str) == CUENTA_EXENTO].copy()
    exento = exento.groupby(["Clave Documento", "Tipo Documento", "Numero Documento", "NIT"], dropna=False)["Movimiento VES"].sum().reset_index().rename(columns={"Movimiento VES": "Exento Contabilidad"})

    retenciones = mayor[mayor["Cuenta Contable"].astype(str) == CUENTA_RETENCION_IVA].copy()
    retenciones = retenciones.groupby(["Clave Documento", "Tipo Documento", "Numero Documento", "NIT"], dropna=False)["Movimiento VES"].sum().reset_index().rename(columns={"Movimiento VES": "Retencion Contabilidad Signo"})
    retenciones["Retencion Contabilidad"] = retenciones["Retencion Contabilidad Signo"].abs()

    resultado = ingresos.merge(base, on=["Clave Documento", "Tipo Documento", "Numero Documento", "NIT"], how="outer")
    resultado = resultado.merge(exento, on=["Clave Documento", "Tipo Documento", "Numero Documento", "NIT"], how="outer")
    resultado = resultado.merge(retenciones, on=["Clave Documento", "Tipo Documento", "Numero Documento", "NIT"], how="outer")

    for columna in ["Ingreso Contabilidad", "Base Contabilidad", "Exento Contabilidad", "Retencion Contabilidad Signo", "Retencion Contabilidad"]:
        if columna in resultado.columns:
            resultado[columna] = resultado[columna].fillna(0)

    return resultado


def construir_libro(auxiliar, resumen_mayor=None):
    libro = auxiliar.copy()
    libro["Ventas Exentas / Exoneradas / No Sujetas"] = np.where(libro["Base Imponible Bs"] == 0, libro["Monto Bs"], 0)
    libro["Base Gravada"] = libro["Base Imponible Bs"]
    libro["Debito Fiscal IVA"] = libro["IVA Bs"]
    libro["IVA Retenido"] = 0.0

    if resumen_mayor is not None:
        ret = resumen_mayor[["Clave Documento", "NIT", "Retencion Contabilidad"]].copy()
        ret = ret.groupby(["Clave Documento", "NIT"], dropna=False)["Retencion Contabilidad"].sum().reset_index()
        libro = libro.merge(ret, on=["Clave Documento", "NIT"], how="left")
        libro["IVA Retenido"] = libro["Retencion Contabilidad"].fillna(0)
        libro.drop(columns=["Retencion Contabilidad"], inplace=True)

    libro["N° Comprobante Retención"] = ""
    libro["Fecha Comprobante Retención"] = ""
    libro["Observaciones"] = ""
    libro.loc[libro["Tipo"].isin(["N/C", "N/D"]), "Observaciones"] = "Documento modificatorio"
    libro.loc[libro["Estado"] == "ANULADA", "Observaciones"] = "DOCUMENTO ANULADO"
    libro.insert(0, "N° Operación", range(1, len(libro) + 1))

    return libro


def construir_conciliacion(libro, resumen):
    columnas = [
        "Clave Documento", "Tipo Documento", "Numero Documento", "NIT",
        "Base Libro", "Base Contabilidad", "Diferencia Base",
        "Exento Libro", "Exento Contabilidad", "Diferencia Exento",
        "IVA Libro", "Retencion Libro", "Retencion Contabilidad",
        "Diferencia Retencion", "Estado"
    ]

    if resumen is None:
        return pd.DataFrame(columns=columnas)

    libro_control = libro[[
        "Clave Documento", "Tipo", "Numero", "Base Gravada",
        "Ventas Exentas / Exoneradas / No Sujetas", "Debito Fiscal IVA", "IVA Retenido"
    ]].copy()

    libro_control.rename(columns={
        "Tipo": "Tipo Documento",
        "Numero": "Numero Documento",
        "Base Gravada": "Base Libro",
        "Ventas Exentas / Exoneradas / No Sujetas": "Exento Libro",
        "Debito Fiscal IVA": "IVA Libro",
        "IVA Retenido": "Retencion Libro",
    }, inplace=True)

    nit_mayor = resumen[["Clave Documento", "NIT", "Base Contabilidad", "Exento Contabilidad", "Retencion Contabilidad"]].copy()
    nit_mayor = nit_mayor.groupby("Clave Documento", as_index=False).agg({
        "NIT": "first",
        "Base Contabilidad": "sum",
        "Exento Contabilidad": "sum",
        "Retencion Contabilidad": "sum",
    })

    conciliacion = libro_control.merge(nit_mayor, on="Clave Documento", how="outer", suffixes=("_Libro", "_Cont"))
    conciliacion["NIT"] = conciliacion["NIT"].fillna("")

    for col in ["Base Libro", "Base Contabilidad", "Exento Libro", "Exento Contabilidad", "IVA Libro", "Retencion Libro", "Retencion Contabilidad"]:
        if col not in conciliacion.columns:
            conciliacion[col] = 0.0
        conciliacion[col] = pd.to_numeric(conciliacion[col], errors="coerce").fillna(0)

    conciliacion["Diferencia Base"] = conciliacion["Base Libro"] - conciliacion["Base Contabilidad"]
    conciliacion["Diferencia Exento"] = conciliacion["Exento Libro"] - conciliacion["Exento Contabilidad"]
    conciliacion["Diferencia Retencion"] = conciliacion["Retencion Libro"] - conciliacion["Retencion Contabilidad"]

    tolerancia = 0.05
    conciliacion["Estado"] = np.where(
        (conciliacion["Diferencia Base"].abs() <= tolerancia) &
        (conciliacion["Diferencia Exento"].abs() <= tolerancia) &
        (conciliacion["Diferencia Retencion"].abs() <= tolerancia),
        "CUADRADO", "DIFERENCIA"
    )

    return conciliacion


def generar_excel(empresa, fecha_inicio, fecha_fin, libro, conciliacion, auxiliar_anulados, mayor):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        columnas_libro = [
            "N° Operación", "Fecha", "Tipo", "Numero", "Factura Afectada", "Cliente", "Nombre",
            "Base Gravada", "Ventas Exentas / Exoneradas / No Sujetas", "Alicuota",
            "Debito Fiscal IVA", "IVA Retenido", "N° Comprobante Retención",
            "Fecha Comprobante Retención", "Monto Bs", "Observaciones",
        ]
        columnas_libro = [c for c in columnas_libro if c in libro.columns]
        libro[columnas_libro].to_excel(writer, sheet_name="LIBRO DE VENTAS", index=False, startrow=5)

        resumen = libro[["Base Gravada", "Ventas Exentas / Exoneradas / No Sujetas", "Debito Fiscal IVA", "IVA Retenido"]].sum().to_frame().T
        resumen.insert(0, "Empresa", empresa)
        resumen.insert(1, "Desde", fecha_inicio)
        resumen.insert(2, "Hasta", fecha_fin)
        resumen.to_excel(writer, sheet_name="RESUMEN IVA", index=False)

        conciliacion.to_excel(writer, sheet_name="CONCILIACIÓN", index=False)
        conciliacion[conciliacion["Estado"] == "DIFERENCIA"].to_excel(writer, sheet_name="DIFERENCIAS", index=False)
        auxiliar_anulados.to_excel(writer, sheet_name="ANULADOS", index=False)

        if mayor is not None:
            mayor.to_excel(writer, sheet_name="DETALLE MAYOR", index=False)

    output.seek(0)
    wb = load_workbook(output)
    thin = Side(style="thin")

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="center")

        ws.freeze_panes = "A7" if ws.title == "LIBRO DE VENTAS" else "A2"

        for column_cells in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column_cells[0].column)
            for cell in column_cells:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except Exception:
                    pass
            ws.column_dimensions[column_letter].width = min(max_length + 2, 35)

        header_row = 6 if ws.title == "LIBRO DE VENTAS" else 1
        for cell in ws[header_row]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(bottom=thin)

        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00'

    ws = wb["LIBRO DE VENTAS"]
    ws["A1"] = "LIBRO DE VENTAS"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = "Empresa"
    ws["B2"] = empresa
    ws["A3"] = "Período desde"
    ws["B3"] = fecha_inicio
    ws["C3"] = "Hasta"
    ws["D3"] = fecha_fin
    ws["A4"] = "Base legal de estructura: Artículo 76 Reglamento Ley IVA"
    ws["A4"].font = Font(italic=True)

    output_final = io.BytesIO()
    wb.save(output_final)
    output_final.seek(0)
    return output_final


# ============================================================
# INTERFAZ DEL MÓDULO (LLAMADA DESDE MAIN.PY)
# ============================================================

def modulo_crear_libro(sucursal):
    st.title(f"📘 Libro de Ventas SENIAT - {sucursal}")
    st.caption("Auxiliar de FA + Mayor Analítico + Conciliación Contable")
    st.divider()

    col1, col2, col3 = st.columns(3)
    with col1:
        empresa = st.text_input("Empresa", value=sucursal)
    with col2:
        fecha_inicio = st.date_input("Fecha inicial")
    with col3:
        fecha_fin = st.date_input("Fecha final")

    with st.expander("⚙️ Configuración de cuentas contables", expanded=False):
        c1, c2 = st.columns(2)
        with c1:
            st.text_input("Cuenta de ingresos", value=CUENTA_INGRESOS, disabled=True)
            st.text_input("Cuenta de base gravada", value=CUENTA_BASE_GRAVADA, disabled=True)
        with c2:
            st.text_input("Cuenta exenta", value=CUENTA_EXENTO, disabled=True)
            st.text_input("Cuenta IVA retenido", value=CUENTA_RETENCION_IVA, disabled=True)

    st.subheader("📂 Archivos de entrada")
    auxiliar_file = st.file_uploader("1. Auxiliar de FA (TXT, CSV, DAT, Excel)", type=["txt", "csv", "xlsx", "xls", "dat"], key=f"aux_{sucursal}")
    mayor_files = st.file_uploader("2. Mayor Analítico - hasta 2 archivos (Opcional)", type=["xlsx", "xls", "csv", "txt"], accept_multiple_files=True, key=f"may_{sucursal}")

    if mayor_files and len(mayor_files) > 2:
        st.error("Solo se permiten hasta 2 archivos de Mayor Analítico.")
        mayor_files = mayor_files[:2]

    st.divider()

    if st.button("🚀 Generar Libro de Ventas y Conciliación", type="primary", use_container_width=True):
        if auxiliar_file is None:
            st.error("Debe cargar el Auxiliar de FA.")
            st.stop()

        if fecha_inicio > fecha_fin:
            st.error("La fecha inicial no puede ser mayor que la fecha final.")
            st.stop()

        try:
            with st.spinner("Procesando Auxiliar de FA..."):
                auxiliar_raw = leer_auxiliar_fa(auxiliar_file)
                auxiliar = preparar_auxiliar_fa(auxiliar_raw)

            auxiliar = auxiliar[
                (auxiliar["Fecha"].dt.date >= fecha_inicio) &
                (auxiliar["Fecha"].dt.date <= fecha_fin)
            ].copy()

            anulados = auxiliar[auxiliar["Estado"] == "ANULADA"].copy()
            activos = auxiliar[auxiliar["Estado"] == "ACTIVA"].copy()

            mayor_total = []
            if mayor_files:
                with st.spinner("Procesando Mayor Analítico..."):
                    for archivo in mayor_files:
                        mayor_raw = leer_mayor(archivo)
                        mayor_preparado = preparar_mayor(mayor_raw)
                        mayor_total.append(mayor_preparado)

            if mayor_total:
                mayor = pd.concat(mayor_total, ignore_index=True)
                mayor = mayor[
                    (mayor["Fecha"].dt.date >= fecha_inicio) &
                    (mayor["Fecha"].dt.date <= fecha_fin)
                ].copy()
                resumen_mayor = resumir_mayor(mayor)
            else:
                mayor = None
                resumen_mayor = None

            with st.spinner("Construyendo Libro de Ventas..."):
                libro = construir_libro(activos, resumen_mayor)

            conciliacion = construir_conciliacion(libro, resumen_mayor)

            st.success("Proceso terminado correctamente.")

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Documentos", len(libro))
            m2.metric("Anulados", len(anulados))
            m3.metric("Cuadrados", (conciliacion["Estado"] == "CUADRADO").sum() if len(conciliacion) > 0 else 0)
            m4.metric("Diferencias", (conciliacion["Estado"] == "DIFERENCIA").sum() if len(conciliacion) > 0 else 0)

            st.subheader("📘 Vista previa - Libro de Ventas")
            cols_preview = ["N° Operación", "Fecha", "Tipo", "Numero", "Factura Afectada", "Cliente", "Nombre", "Base Gravada", "Ventas Exentas / Exoneradas / No Sujetas", "Alicuota", "Debito Fiscal IVA", "IVA Retenido", "Monto Bs"]
            cols_preview = [c for c in cols_preview if c in libro.columns]
            st.dataframe(libro[cols_preview], use_container_width=True, hide_index=True)

            if resumen_mayor is not None:
                st.subheader("🔎 Conciliación con Contabilidad")
                st.dataframe(conciliacion, use_container_width=True, hide_index=True)

            with st.spinner("Generando Excel final..."):
                archivo_excel = generar_excel(
                    empresa=empresa,
                    fecha_inicio=fecha_inicio,
                    fecha_fin=fecha_fin,
                    libro=libro,
                    conciliacion=conciliacion,
                    auxiliar_anulados=anulados,
                    mayor=mayor,
                )

            st.download_button(
                label="📥 Descargar Libro de Ventas en Excel",
                data=archivo_excel,
                file_name=f"Libro_Ventas_{empresa.replace(',', '').replace(' ', '_')}_{fecha_inicio}_{fecha_fin}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error("No fue posible procesar los archivos.")
            st.exception(e)
