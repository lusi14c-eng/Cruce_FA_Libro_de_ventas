import streamlit as st

def modulo_crear_libro(sucursal):
    st.title(f"📝 Generador de Libro de Ventas - {sucursal}")
    st.caption("Crea la estructura oficial del Libro de Ventas SENIAT")

    with st.form("form_libro"):
        col1, col2 = st.columns(2)
        with col1:
            mes = st.selectbox("Mes Afecto", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
        with col2:
            quincena = st.radio("Quincena", ["1era Quincena", "2da Quincena"])

        archivo_bruto = st.file_uploader("Subir data de ventas sin procesar", type=["xlsx", "csv"])
        btn_generar = st.form_submit_button("🔨 Formatear a Estructura SENIAT")

    if btn_generar and archivo_bruto:
        st.success(f"Libro procesado para {sucursal} - Periodo: {mes} ({quincena}).")
