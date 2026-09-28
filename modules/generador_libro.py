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
CUENTA_BASE_GRAVADA = "2.1.3.04.1.001"  # Pasivo IVA (Opcional)
CUENTA_RETENCION_IVA = "2.1.3.04.1.006"  # Retenciones IVA (Opcional)


# ============================================================
# FUNCIONES GENERALES Y LIMPIEZA
# ============================================================


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor).strip()
    valor = (
        unicodedata.normalize("NFKD", valor)
        .encode("ascii", "ignore")
        .decode("ascii")
    )
    return valor.upper().strip()


def limpiar_numero(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float, np.number)):
        return float(valor)
    texto = str(valor).strip()
    if texto == "":
        return 0.0

    texto = (
        texto.replace("$", "")
        .replace("Bs.", "")
        .replace("Bs", "")
        .replace(" ", "")
    )

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
    if texto in [
        "N/C",
        "NC",
        "NOTA DE CREDITO",
        "NOTA CREDITO",
        "DEV",
        "DEV#",
        "DEVOLUCION",
    ]:
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
        r"^(DEV)#?(.*)$",
        r"^(DEVOLUCION)#?(.*)$",
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
                on_bad_lines="skip",
            )
        except Exception:
            df = pd.read_csv(
                io.StringIO(texto),
                sep="|",
                dtype=str,
                engine="c",
                on_bad_lines="skip",
            )
    elif nombre.endswith((".xlsx", ".xls")):
        df = pd.read_excel(io.BytesIO(contenido), dtype=str)
    else:
        raise ValueError("Formato de Auxiliar FA no soportado.")

    df = df.dropna(axis=1, how="all")
    df.columns = [str(col).strip().strip('"').strip("'") for col in df.columns]
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = (
            df[col].astype(str).str.replace('"', "", regex=False).str.strip()
        )

    return df


def preparar_auxiliar_fa(df):
    df = df.copy()

    mapa = {
        "numero": detectar_columna(
            df,
            [
                "Numero",
                "Número",
                "Nro",
                "Nro.",
                "Num",
                "Num.",
                "Documento",
                "DOCUMENTO",
                "Nro Documento",
                "Nro. Documento",
                "Nro Doc",
                "Nro. Doc",
                "FACTURA",
                "DOC_NUM",
            ],
        ),
        "tipo": detectar_columna(
            df, ["Tipo", "Tipo Doc", "Tipo Documento", "TIP_DOC", "TIPO_DOC"]
        ),
        "fecha": detectar_columna(
            df,
            [
                "Fecha Doc",
                "Fecha",
                "Fecha Documento",
                "FEC_DOC",
                "FECHA_DOC",
            ],
        ),
        "anulada": detectar_columna(
            df,
            ["ANULADA", "Anulada", "Anulado", "Status", "Estado", "ANULADO"],
        ),
        "moneda": detectar_columna(df, ["moneda_factura", "Moneda", "MONEDA"]),
        "subtipo": detectar_columna(df, ["Sub Tipo", "Subtipo", "SUB_TIPO"]),
        "factura_afectada": detectar_columna(
            df,
            [
                "Factura Afectada",
                "Doc Afectado",
                "Afectado",
                "FACTURA_AFECTADA",
                "U_NUM_FAC_AFECTADA",
            ],
        ),
        "cliente": detectar_columna(
            df,
            [
                "RIF",
                "Nit",
                "NIT",
                "Cliente",
                "Cod Cliente",
                "Código Cliente",
                "COD_CLI",
            ],
        ),
        "nombre": detectar_columna(
            df,
            [
                "NOMBRE_CLIENTE",
                "Nombre Cliente",
                "Nombre",
                "Razon Social",
                "Razón Social",
                "NOM_CLI",
            ],
        ),
        "asiento": detectar_columna(df, ["Asiento", "Nro Asiento"]),
        "modulo": detectar_columna(df, ["Modulo", "Módulo"]),
        "tipo_cambio": detectar_columna(
            df, ["Tipo Cambio", "Tasa", "Tasa Cambio"]
        ),
        "costo_total_bs": detectar_columna(
            df, ["Costo Total Bs.", "Costo Total Bs", "Costo Total"]
        ),
        "subtotal_bs": detectar_columna(
            df,
            [
                "SUBTOTAL Bs.",
                "SUBTOTAL Bs",
                "Subtotal Bs.",
                "Subtotal Bs",
                "SUBTOTAL",
            ],
        ),
        "base_bs": detectar_columna(
            df,
            [
                "Base Imponible Bs.",
                "Base Imponible Bs",
                "Base Imponible",
                "Base Bs.",
                "Base Bs",
                "Base Gravada",
                "Monto Base",
                "BASE_IMP",
                "M_BASE",
                "BASE_IMP_IVA_Bs.",
                "BASE_IMP_IVA_Bs",
                "BASE_IMP_IVA Bs.",
            ],
        ),
        "iva_bs": detectar_columna(
            df,
            [
                "Iva Bs.",
                "IVA Bs.",
                "Iva Bs",
                "IVA Bs",
                "Monto IVA",
                "IVA",
                "IVA Bs.",
            ],
        ),
        "flete_bs": detectar_columna(
            df, ["FLETES Bs.", "FLETES Bs", "Flete Bs.", "Flete Bs", "Flete"]
        ),
        "monto_bs": detectar_columna(
            df,
            [
                "Monto Bs.",
                "Monto Bs",
                "Total Bs.",
                "Total Bs",
                "Total",
                "Monto Total",
                "MONTO Bs.",
            ],
        ),
        "base_usd": detectar_columna(
            df, ["Base Imponible $", "Base Imponible USD", "Base $", "BASE_IMP_IVA $"]
        ),
        "iva_usd": detectar_columna(df, ["Iva $", "IVA $", "IVA USD"]),
        "monto_usd": detectar_columna(
            df, ["Monto $", "Monto USD", "Total $"]
        ),
    }
    obligatorias = [
        "numero",
        "tipo",
        "fecha",
        "nombre",
        "base_bs",
        "iva_bs",
        "monto_bs",
    ]
    faltantes = [campo for campo in obligatorias if mapa[campo] is None]

    if faltantes:
        raise ValueError(
            f"No se encontraron las columnas necesarias del Auxiliar FA: {faltantes}"
        )

    salida = pd.DataFrame()
    salida["Numero"] = df[mapa["numero"]].apply(limpiar_documento)
    salida["Tipo"] = df[mapa["tipo"]].apply(tipo_documento)
    salida["Fecha"] = pd.to_datetime(df[mapa["fecha"]], errors="coerce")
    salida["Anulada"] = (
        df[mapa["anulada"]].astype(str).str.upper().str.strip()
        if mapa["anulada"]
        else ""
    )
    salida["Moneda"] = df[mapa["moneda"]].astype(str) if mapa["moneda"] else ""
    salida["Sub Tipo"] = (
        df[mapa["subtipo"]].astype(str) if mapa["subtipo"] else ""
    )
    salida["Factura Afectada"] = (
        df[mapa["factura_afectada"]].apply(limpiar_documento)
        if mapa["factura_afectada"]
        else ""
    )
    salida["Cliente"] = (
        df[mapa["cliente"]].astype(str) if mapa["cliente"] else ""
    )
    salida["Nombre"] = df[mapa["nombre"]].astype(str)
    salida["Asiento"] = (
        df[mapa["asiento"]].astype(str) if mapa["asiento"] else ""
    )
    salida["Modulo"] = df[mapa["modulo"]].astype(str) if mapa["modulo"] else ""
    salida["Tipo Cambio"] = (
        df[mapa["tipo_cambio"]].apply(limpiar_numero)
        if mapa["tipo_cambio"]
        else 0
    )
    salida["Costo Total Bs"] = (
        df[mapa["costo_total_bs"]].apply(limpiar_numero)
        if mapa["costo_total_bs"]
        else 0
    )
    salida["Subtotal Bs"] = (
        df[mapa["subtotal_bs"]].apply(limpiar_numero)
        if mapa["subtotal_bs"]
        else 0
    )
    salida["Base Imponible Bs"] = df[mapa["base_bs"]].apply(limpiar_numero)
    salida["IVA Bs"] = df[mapa["iva_bs"]].apply(limpiar_numero)
    salida["Flete Bs"] = (
        df[mapa["flete_bs"]].apply(limpiar_numero) if mapa["flete_bs"] else 0
    )
    salida["Monto Bs"] = df[mapa["monto_bs"]].apply(limpiar_numero)
    salida["Base Imponible $"] = (
        df[mapa["base_usd"]].apply(limpiar_numero) if mapa["base_usd"] else 0
    )
    salida["IVA $"] = (
        df[mapa["iva_usd"]].apply(limpiar_numero) if mapa["iva_usd"] else 0
    )
    salida["Monto $"] = (
        df[mapa["monto_usd"]].apply(limpiar_numero) if mapa["monto_usd"] else 0
    )

    # N/C Signo Negativo
    es_nc = salida["Tipo"] == "N/C"
    cols_a_negativizar = [
        "Costo Total Bs",
        "Subtotal Bs",
        "Base Imponible Bs",
        "IVA Bs",
        "Flete Bs",
        "Monto Bs",
        "Base Imponible $",
        "IVA $",
        "Monto $",
    ]

    for col in cols_a_negativizar:
        val_abs = salida[col].abs()
        salida[col] = np.where(es_nc, -val_abs, val_abs)

    salida["Clave Documento"] = (
        salida["Tipo"].astype(str) + "|" + salida["Numero"].astype(str)
    )
    salida["Alicuota"] = np.where(
        salida["Base Imponible Bs"] != 0,
        (salida["IVA Bs"] / salida["Base Imponible Bs"] * 100),
        0,
    )
    salida["Alicuota"] = salida["Alicuota"].round(2)
    salida["Estado"] = np.where(
        salida["Anulada"].isin(["S", "SI", "Y", "YES"]), "ANULADA", "ACTIVA"
    )

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
            df = pd.read_csv(
                io.BytesIO(contenido), sep="|", dtype=str, engine="python"
            )
        except Exception:
            df = pd.read_csv(
                io.BytesIO(contenido), sep="\t", dtype=str, engine="python"
            )
    else:
        raise ValueError("Formato de mayor no soportado.")

    df = df.dropna(axis=1, how="all")
    df.columns = [str(c).strip().strip('"').strip("'") for c in df.columns]
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = (
            df[col].astype(str).str.replace('"', "", regex=False).str.strip()
        )
    return df


def preparar_mayor(df):
    df = df.copy()

    col_asiento = detectar_columna(df, ["Asiento"])
    col_fecha = detectar_columna(df, ["Fecha"])
    col_cuenta = detectar_columna(df, ["Cuenta Contable"])
    col_descripcion = detectar_columna(
        df, ["Descripción de la cuenta", "Descripcion de la cuenta"]
    )
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

    faltantes = [
        nombre for nombre, columna in obligatorias.items() if columna is None
    ]
    if faltantes:
        raise ValueError(
            "El mayor no contiene las columnas requeridas: "
            + ", ".join(faltantes)
        )

    salida = pd.DataFrame()
    salida["Asiento"] = df[col_asiento].astype(str)
    salida["Fecha"] = pd.to_datetime(df[col_fecha], errors="coerce")
    salida["Cuenta Contable"] = df[col_cuenta].astype(str).str.strip()
    salida["Descripción de la cuenta"] = (
        df[col_descripcion].astype(str) if col_descripcion else ""
    )
    salida["Fuente"] = df[col_fuente].astype(str).str.strip()
    salida["Referencia"] = (
        df[col_referencia].astype(str).str.strip() if col_referencia else ""
    )
    salida["Débito VES"] = df[col_debito].apply(limpiar_numero)
    salida["Crédito VES"] = df[col_credito].apply(limpiar_numero)
    salida["NIT"] = df[col_nit].apply(limpiar_rif)

    # NETO DE INGRESO: Créditos menos Débitos
    salida["Movimiento VES"] = salida["Crédito VES"] - salida["Débito VES"]

    parsed = salida["Fuente"].apply(separar_fuente)
    salida["Tipo Documento"] = parsed.apply(lambda x: x[0])
    salida["Numero Documento"] = parsed.apply(
        lambda x: limpiar_documento(x[1])
    )
    salida["Clave Documento"] = (
        salida["Tipo Documento"] + "|" + salida["Numero Documento"]
    )

    return salida


def resumir_mayor(mayor):
    # Cuentas de Ingreso: Todas las que empiecen por 4.1.1 + la cuenta exenta 7.1.3.45.1.997
    es_ingreso_total = (
        mayor["Cuenta Contable"].str.startswith(CUENTA_INGRESOS)
    ) | (mayor["Cuenta Contable"].astype(str) == CUENTA_EXENTO)

    ingresos = mayor[es_ingreso_total].copy()
    total_ingreso = (
        ingresos.groupby(
            ["Clave Documento", "Tipo Documento", "Numero Documento", "NIT"],
            dropna=False,
        )["Movimiento VES"]
        .sum()
        .reset_index()
        .rename(columns={"Movimiento VES": "Total Ingreso Contabilidad"})
    )

    # Pasivo IVA (2.1.3.04.1.001) - Opcional
    es_iva = mayor["Cuenta Contable"].astype(str) == CUENTA_BASE_GRAVADA
    iva = mayor[es_iva].copy()
    if not iva.empty:
        iva = (
            iva.groupby(
                ["Clave Documento", "Tipo Documento", "Numero Documento", "NIT"],
                dropna=False,
            )["Movimiento VES"]
            .sum()
            .reset_index()
            .rename(columns={"Movimiento VES": "IVA Contabilidad"})
        )
    else:
        iva = pd.DataFrame(
            columns=[
                "Clave Documento",
                "Tipo Documento",
                "Numero Documento",
                "NIT",
                "IVA Contabilidad",
            ]
        )

    # Retenciones IVA (2.1.3.04.1.006) - Opcional
    es_retencion = mayor["Cuenta Contable"].astype(str) == CUENTA_RETENCION_IVA
    retenciones = mayor[es_retencion].copy()
    if not retenciones.empty:
        retenciones = (
            retenciones.groupby(
                ["Clave Documento", "Tipo Documento", "Numero Documento", "NIT"],
                dropna=False,
            )["Movimiento VES"]
            .sum()
            .reset_index()
            .rename(columns={"Movimiento VES": "Retencion Contabilidad Signo"})
        )
        retenciones["Retencion Contabilidad"] = retenciones[
            "Retencion Contabilidad Signo"
        ].abs()
    else:
        retenciones = pd.DataFrame(
            columns=[
                "Clave Documento",
                "Tipo Documento",
                "Numero Documento",
                "NIT",
                "Retencion Contabilidad",
            ]
        )

    resultado = total_ingreso.merge(
        iva,
        on=["Clave Documento", "Tipo Documento", "Numero Documento", "NIT"],
        how="outer",
    )
    resultado = resultado.merge(
        retenciones,
        on=["Clave Documento", "Tipo Documento", "Numero Documento", "NIT"],
        how="outer",
    )

    cols_num = [
        "Total Ingreso Contabilidad",
        "IVA Contabilidad",
        "Retencion Contabilidad",
    ]
    for col in cols_num:
        if col in resultado.columns:
            resultado[col] = resultado[col].fillna(0.0)
        else:
            resultado[col] = 0.0

    return resultado


def construir_libro(auxiliar, resumen_mayor=None):
    libro = auxiliar.copy()

    # Exento = Subtotal - Base Imponible + Fletes
    libro["Ventas Exentas / Exoneradas / No Sujetas"] = (
        libro["Subtotal Bs"] - libro["Base Imponible Bs"] + libro["Flete Bs"]
    )
    libro["Ventas Exentas / Exoneradas / No Sujetas"] = libro[
        "Ventas Exentas / Exoneradas / No Sujetas"
    ].round(2)

    libro["Base Gravada"] = libro["Base Imponible Bs"]
    libro["Debito Fiscal IVA"] = libro["IVA Bs"]
    libro["IVA Retenido"] = 0.0

    if resumen_mayor is not None:
        rif_map = resumen_mayor[["Clave Documento", "NIT"]].drop_duplicates(
            subset=["Clave Documento"]
        )
        libro = libro.merge(rif_map, on="Clave Documento", how="left")
        if "NIT" in libro.columns:
            libro["Cliente"] = np.where(
                libro["NIT"].fillna("") != "", libro["NIT"], libro["Cliente"]
            )
            libro.drop(columns=["NIT"], inplace=True)

        if "Retencion Contabilidad" in resumen_mayor.columns:
            ret = resumen_mayor[
                ["Clave Documento", "Retencion Contabilidad"]
            ].copy()
            ret = (
                ret.groupby("Clave Documento", dropna=False)[
                    "Retencion Contabilidad"
                ]
                .sum()
                .reset_index()
            )
            libro = libro.merge(ret, on="Clave Documento", how="left")
            libro["IVA Retenido"] = libro["Retencion Contabilidad"].fillna(0)
            if "Retencion Contabilidad" in libro.columns:
                libro.drop(columns=["Retencion Contabilidad"], inplace=True)

    libro["Fecha"] = pd.to_datetime(libro["Fecha"]).dt.strftime("%Y-%m-%d")
    libro["N° Comprobante Retención"] = ""
    libro["Fecha Comprobante Retención"] = ""
    libro["Observaciones"] = ""
    libro.loc[
        libro["Tipo"].isin(["N/C", "N/D"]), "Observaciones"
    ] = "Documento modificatorio"
    libro.loc[libro["Estado"] == "ANULADA", "Observaciones"] = (
        "DOCUMENTO ANULADO"
    )
    libro.insert(0, "N° Operación", range(1, len(libro) + 1))

    return libro


def construir_conciliacion(libro, resumen):
    columnas = [
        "Clave Documento",
        "Tipo Documento",
        "Numero Documento",
        "NIT",
        "Base Libro",
        "Exento Libro",
        "Total Ingreso Libro",
        "Total Ingreso Contabilidad",
        "Diferencia Total Ingreso",
        "IVA Libro",
        "IVA Contabilidad",
        "Retencion Libro",
        "Retencion Contabilidad",
        "Estado",
    ]

    if resumen is None or resumen.empty:
        return pd.DataFrame(columns=columnas)

    libro_control = libro[
        [
            "Clave Documento",
            "Tipo",
            "Numero",
            "Base Gravada",
            "Ventas Exentas / Exoneradas / No Sujetas",
            "Debito Fiscal IVA",
            "IVA Retenido",
        ]
    ].copy()

    libro_control.rename(
        columns={
            "Tipo": "Tipo Documento",
            "Numero": "Numero Documento",
            "Base Gravada": "Base Libro",
            "Ventas Exentas / Exoneradas / No Sujetas": "Exento Libro",
            "Debito Fiscal IVA": "IVA Libro",
            "IVA Retenido": "Retencion Libro",
        },
        inplace=True,
    )

    nit_mayor = resumen[
        [
            "Clave Documento",
            "Tipo Documento",
            "Numero Documento",
            "NIT",
            "Total Ingreso Contabilidad",
            "IVA Contabilidad",
            "Retencion Contabilidad",
        ]
    ].copy()

    nit_mayor = nit_mayor.groupby("Clave Documento", as_index=False).agg({
        "Tipo Documento": "first",
        "Numero Documento": "first",
        "NIT": "first",
        "Total Ingreso Contabilidad": "sum",
        "IVA Contabilidad": "sum",
        "Retencion Contabilidad": "sum",
    })

    conciliacion = libro_control.merge(
        nit_mayor, on="Clave Documento", how="outer", suffixes=("_libro", "_mayor")
    )
    
    # Consolidar metadatos tras el outer join
    conciliacion["Tipo Documento"] = conciliacion["Tipo Documento_libro"].combine_first(conciliacion["Tipo Documento_mayor"])
    conciliacion["Numero Documento"] = conciliacion["Numero Documento_libro"].combine_first(conciliacion["Numero Documento_mayor"])
    conciliacion.drop(columns=["Tipo Documento_libro", "Tipo Documento_mayor", "Numero Documento_libro", "Numero Documento_mayor"], inplace=True)

    conciliacion["NIT"] = conciliacion["NIT"].fillna("")

    cols_eval = [
        "Base Libro",
        "Exento Libro",
        "Total Ingreso Contabilidad",
        "IVA Libro",
        "IVA Contabilidad",
        "Retencion Libro",
        "Retencion Contabilidad",
    ]
    for col in cols_eval:
        if col not in conciliacion.columns:
            conciliacion[col] = 0.0
        conciliacion[col] = pd.to_numeric(
            conciliacion[col], errors="coerce"
        ).fillna(0.0)

    # SUMA GLOBAL DE INGRESO LIBRO
    conciliacion["Total Ingreso Libro"] = (
        conciliacion["Base Libro"] + conciliacion["Exento Libro"]
    )

    # DIFERENCIA DE INGRESO
    conciliacion["Diferencia Total Ingreso"] = (
        conciliacion["Total Ingreso Libro"]
        - conciliacion["Total Ingreso Contabilidad"]
    )

    tolerancia = 0.05

    tiene_iva_conta = (conciliacion["IVA Contabilidad"].abs() > 0).any()

    if tiene_iva_conta:
        condicion_cuadre = (
            (conciliacion["Diferencia Total Ingreso"].abs() <= tolerancia)
            & (
                (conciliacion["IVA Libro"] - conciliacion["IVA Contabilidad"]).abs()
                <= tolerancia
            )
        )
    else:
        condicion_cuadre = (
            conciliacion["Diferencia Total Ingreso"].abs() <= tolerancia
        )

    conciliacion["Estado"] = np.where(condicion_cuadre, "CUADRADO", "DIFERENCIA")

    return conciliacion[columnas]
