import streamlit as st
import pandas as pd
from utils import normalizar_turno, encontrar_mejor_coincidencia, formatear_rut

st.set_page_config(page_title="BUK Shift Importer", layout="wide")

st.title("🛠️ Generador de Importador BUK")
st.markdown("Transforma matrices de turnos manuales al formato de carga masiva de BUK.")

uploaded_file = st.file_uploader("Carga el Excel con las 3 hojas", type=["xlsx"])

if uploaded_file:
    try:
        # Cargar hojas
        df_matriz = pd.read_excel(uploaded_file, sheet_name=0)
        df_colab = pd.read_excel(uploaded_file, sheet_name=1)
        df_cat = pd.read_excel(uploaded_file, sheet_name=2)

        st.info("🔄 Procesando datos y cruzando nombres...")

        # 1. Normalizar Catálogo (Hoja 3)
        # Col 0: Sigla, Col 1: Horario
        df_cat['horario_norm'] = df_cat.iloc[:, 1].apply(normalizar_turno)
        mapa_turnos = dict(zip(df_cat['horario_norm'], df_cat.iloc[:, 0]))

        # 2. Normalizar Base de Colaboradores (Hoja 2)
        df_colab.columns = [c.strip() for c in df_colab.columns]
        # Aseguramos nombres en mayúsculas para el match
        df_colab['Nombre_Upper'] = df_colab.iloc[:, 0].str.upper()

        # 3. Procesar Matriz (Hoja 1)
        # Identificar columna de nombres (la primera)
        col_nombres_matriz = df_matriz.columns[0]
        
        # Crear mapeo de nombres usando Fuzzy Matching
        nombres_completos_lista = df_colab['Nombre_Upper'].tolist()
        mapping_nombres = {n: encontrar_mejor_coincidencia(n, nombres_completos_lista) 
                          for n in df_matriz[col_nombres_matriz].unique()}

        # 4. Construcción del DataFrame Final
        df_final = df_matriz.copy()
        df_final['Nombre Completo'] = df_final[col_nombres_matriz].map(mapping_nombres)
        
        # Traducir horarios a Siglas en las columnas de fechas
        columnas_fechas = df_matriz.columns[1:]
        for col in columnas_fechas:
            df_final[col] = df_final[col].apply(normalizar_turno).map(mapa_turnos).fillna(df_final[col])

        # 5. Unir con Info de Colaborador (RUT, Área, Supervisor)
        # Traemos las primeras 4 columnas de la Hoja 2
        info_colab = df_colab.iloc[:, [0, 1, 2, 3]]
        info_colab.columns = ['Nombre Completo', 'RUT', 'Área', 'Supervisor']
        
        # Merge Final
        resultado = pd.merge(info_colab, df_final, on='Nombre Completo', how='inner')
        
        # Limpieza final: quitar columna original de nombre corto
        resultado = resultado.drop(columns=[col_nombres_matriz])

        st.success("✅ ¡Cruce completado!")
        st.dataframe(resultado.head(10))

        # Botón de Descarga
        csv = resultado.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 Descargar CSV para BUK",
            data=csv,
            file_name="importador_buk_final.csv",
            mime="text/csv",
        )

    except Exception as e:
        st.error(f"Hubo un error al procesar el archivo: {e}")

else:
    st.warning("Esperando archivo Excel...")
