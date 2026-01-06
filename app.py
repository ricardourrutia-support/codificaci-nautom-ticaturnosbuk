import streamlit as st
import pandas as pd
import io
import utils  # Importación directa del módulo

st.set_page_config(page_title="BUK Shift Importer PRO", layout="wide")

st.title("🚀 Generador de Importador BUK")

uploaded_file = st.file_uploader("Carga el Excel 'Turnos Formato Supervisor'", type=["xlsx"])

if uploaded_file:
    try:
        # 1. Cargar Hojas
        df_raw_turnos = pd.read_excel(uploaded_file, sheet_name=0, header=None)
        df_colab = pd.read_excel(uploaded_file, sheet_name=1)
        df_cat = pd.read_excel(uploaded_file, sheet_name=2)

        # 2. Procesar Fechas desde la Fila 2
        row_fechas = df_raw_turnos.iloc[1, 1:]
        fechas_columnas = []
        for f in row_fechas:
            try:
                if pd.isna(f): fechas_columnas.append("VACIO")
                else: fechas_columnas.append(pd.to_datetime(f).strftime('%d-%m-%Y'))
            except:
                fechas_columnas.append(str(f).strip())

        # 3. Preparar Matriz de Turnos
        df_turnos = df_raw_turnos.iloc[2:].copy()
        df_turnos.columns = ['Nombre_Corto'] + fechas_columnas
        df_turnos = df_turnos[df_turnos['Nombre_Corto'].notna()]

        # 4. Preparar Catálogo de Turnos (Usando utils.normalizar_turno)
        df_cat['horario_norm'] = df_cat.iloc[:, 1].apply(utils.normalizar_turno)
        mapa_turnos = dict(zip(df_cat['horario_norm'], df_cat.iloc[:, 0]))
        mapa_turnos["L"] = "L"

        # 5. Cruzar Nombres con Fuzzy Matching
        nombres_base = df_colab.iloc[:, 0].astype(str).str.upper().tolist()
        
        # Barra de progreso para el match de nombres
        nombres_unicos = df_turnos['Nombre_Corto'].unique()
        dict_nombres = {}
        progreso = st.progress(0)
        for i, n in enumerate(nombres_unicos):
            dict_nombres[n] = utils.encontrar_mejor_coincidencia(n, nombres_base)
            progreso.progress((i + 1) / len(nombres_unicos))
        
        df_turnos['Nombre Completo'] = df_turnos['Nombre_Corto'].map(dict_nombres)

        # 6. Traducir horarios a Siglas
        cols_fechas = [c for c in fechas_columnas if c != "VACIO"]
        for col in cols_fechas:
            df_turnos[col] = df_turnos[col].apply(utils.normalizar_turno).map(mapa_turnos).fillna(df_turnos[col])

        # 7. Unión Final
        info_colab = df_colab.iloc[:, [0, 1, 2, 3]].copy()
        info_colab.columns = ['Nombre Completo', 'RUT', 'Área', 'Supervisor']
        info_colab = info_colab.drop_duplicates(subset=['RUT'])

        resultado = pd.merge(info_colab, df_turnos[['Nombre Completo'] + cols_fechas], on='Nombre Completo', how='inner')

        st.success(f"✅ Procesados {len(resultado)} colaboradores.")
        st.dataframe(resultado)

        # 8. Exportar a Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            resultado.to_excel(writer, index=False, sheet_name='BUK')
        
        st.download_button(
            label="📥 Descargar Excel para BUK",
            data=output.getvalue(),
            file_name="Importador_BUK.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error detectado: {e}")
