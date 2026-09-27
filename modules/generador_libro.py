import csv

def leer_auxiliar_fa(uploaded_file):
    nombre = uploaded_file.name.lower()
    contenido = uploaded_file.read()

    if nombre.endswith((".txt", ".csv", ".dat")):
        texto = contenido.decode("utf-8", errors="replace")
        
        # 1. Intentar lectura ignorando comillas como caracteres de control (quoting=3)
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
            # 2. Respaldo estándar de lectura por líneas
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
    # Limpiar comillas sobrantes de los nombres de columnas
    df.columns = [str(col).strip().strip('"') for col in df.columns]
    
    # Limpiar comillas sobrantes en los datos de texto
    df = df.apply(lambda col: col.str.strip('"') if col.dtype == "object" else col)
    
    return df
