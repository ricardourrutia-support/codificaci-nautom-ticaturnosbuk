import streamlit as st
import pandas as pd
import io
from utils import normalizar_turno, encontrar_mejor_coincidencia

st.set_page_config(page_title="BUK Shift Importer PRO", layout="wide")

st.title("🚀 Generador de Importador BUK")

uploaded_file = st.file_uploader("Carga el Excel 'Turnos Formato Supervisor'", type=["xlsx"])

if uploaded_file:
    # 1. Lectura de datos
    df_raw_turnos = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    df_colab = pd.read_excel(uploaded_file, sheet_name=1)
    df_cat = pd.read_excel(uploaded_file, sheet_name=2)

    # --- PROCESAMIENTO DE FECHAS (Hoja 1) ---
    # Extraemos las fechas de la fila 2 (índice 1 en Python)
    fechas = df_raw_turnos.iloc[1, 1:].tolist()
    # Convertimos a formato DD-MM-YYYY
    fechas_limpias = []
    for f in fechas:
        try:
            fechas_limpias.append(pd.to_datetime(f).strftime('%d-%m-%Y'))
        except:
            fechas_limpias.append(str(f))

    # Reestructuramos la matriz
    df_turnos = df_raw_turnos.iloc[2:].copy()
    df_turnos.columns = ['Colaborador_Original'] + fechas_limpias

    # --- NORMALIZACIÓN DEL CATÁLOGO ---
    df_cat['horario_norm'] = df_cat.iloc[:, 1].apply(normalizar_turno)
    mapa_turnos = dict(zip(df_cat['horario_norm'], df_cat.iloc[:, 0]))
    mapa_turnos["L"] = "L" # Asegurar que Libre -> L

    # --- CRUCE DE COLABORADORES ---
    nombres_base = df_colab.iloc[:, 0].str.upper().tolist()
    
    # Mapear nombres cortos a nombres completos
    nombres_unicos_matriz = df_turnos['Colaborador_Original'].unique()
    dict_nombres = {n: encontrar_mejor_coincidencia(n, nombres_base) for n in nombres_unicos_matriz}
    
    df_turnos['Nombre Completo'] = df_turnos['Colaborador_Original'].map(dict_nombres)

    # --- TRADUCCIÓN DE TURNOS A SIGLAS ---
    for col in fechas_limpias:
        df_turnos[col] = df_turnos[col].apply(normalizar_turno).map(mapa_turnos).fillna(df_turnos[col])

    # --- MERGE FINAL ---
    info_colab = df_colab.iloc[:, [0, 1, 2, 3]]
    info_colab.columns = ['Nombre Completo', 'RUT', 'Área', 'Supervisor']
    
    # Eliminar duplicados en la base de colaboradores para evitar filas extra
    info_colab = info_colab.drop_duplicates(subset=['Nombre Completo'])
    
    resultado = pd.merge(info_colab, df_turnos, on='Nombre Completo', how='inner')
    resultado = resultado.drop(columns=['Colaborador_Original'])

    st.success(f"✅ Procesados {len(resultado)} colaboradores.")
    st.dataframe(resultado.head(10))

    # --- EXPORTACIÓN A EXCEL ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        resultado.to_excel(writer, index=False, sheet_name='Importador_BUK')
    
    st.download_button(
        label="📥 Descargar Excel para BUK",
        data=output.getvalue(),
        file_name="Importador_Turnos_BUK.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
