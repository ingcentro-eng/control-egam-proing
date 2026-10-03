import streamlit as st
import pandas as pd
import requests
import io
import time

# URL de tu archivo Excel en GitHub (Asegúrate de usar la URL 'RAW')
# Ejemplo: https://raw.githubusercontent.com/TU_USUARIO/TU_REPOSO/main/Base_DP.xlsx
URL_EXCEL_GITHUB = "https://raw.githubusercontent.com/TU_USUARIO/TU_REPOSO/main/Base_DP.xlsx"

@st.cache_data(ttl=60) # Refresca automáticamente el caché cada 60 segundos
def cargar_datos_desde_github(url):
    # Agregamos un parámetro de timestamp para evitar que el navegador responda con caché antiguo
    url_dinamica = f"{url}?v={int(time.time())}"
    respuesta = requests.get(url_dinamica)
    if respuesta.status_code == 200:
        return io.BytesIO(respuesta.content)
    else:
        return None

# Intentar cargar desde GitHub o por File Uploader
archivo_subido = st.file_uploader("📂 Cargar reporte del día (Opcional)", type=["xlsx", "xls"])

if archivo_subido is not None:
    archivo_a_usar = archivo_subido
else:
    archivo_a_usar = cargar_datos_desde_github(URL_EXCEL_GITHUB)
