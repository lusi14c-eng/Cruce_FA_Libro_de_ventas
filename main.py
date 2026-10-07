import streamlit as st
from modules.conciliacion import modulo_conciliacion
from modules.generador_libro import modulo_crear_libro
from modules.logic_analizador_variaciones_gastos import modulo_analizador_gastos

st.set_page_config(
    page_title="Sistema de Conciliación y Libros de Ventas",
    page_icon="📊",
    layout="wide",
)

st.sidebar.title("🏢 Control Central")

sucursal = st.sidebar.selectbox(
    "Selecciona la Sucursal:",
    ["Sucursal Principal (Sillaca)", "Sucursal 2", "Sucursal 3"]
)

st.sidebar.markdown("---")

modulo = st.sidebar.radio(
    "Selecciona el Módulo:",
    [
        "📊 Conciliación de Ventas",
        "📝 Crear Libro de Ventas SENIAT",
        "📊 Analizador de Variaciones de Gastos"
    ]
)

st.caption(f"📍 Operando en: **{sucursal}**")

if modulo == "📊 Conciliación de Ventas":
    modulo_conciliacion(sucursal)
elif modulo == "📝 Crear Libro de Ventas SENIAT":
    modulo_crear_libro(sucursal)
elif modulo == "📊 Analizador de Variaciones de Gastos":
    modulo_analizador_gastos(sucursal)
