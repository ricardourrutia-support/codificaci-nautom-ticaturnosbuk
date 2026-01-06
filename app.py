import streamlit as st
import pandas as pd
import io
from utils import normalizar_turno, encontrar_mejor_coincidencia

st.set_page_config(page_title="BUK Shift Importer PRO", layout="wide")

st.title("🛠️ Generador de Importador BUK")
st.markdown("""
Esta aplicación procesa la matriz de turnos, cruza con la base de colaboradores 
y traduce los horarios a siglas oficiales para la carga masiva en BUK.
""")

# Selector de archivo
uploaded_file = st.file_uploader("Carga el Excel 'Turnos Formato Supervisor'", type=["xlsx"])

if uploaded_file:
    try:
        # 1. LECTURA DE DATOS
        # Cargamos las 3 hojas. Usamos header=None para la matriz porque las fechas están en la fila 2
        df_raw_turnos = pd.read_excel(uploaded_file, sheet_name=0, header=None)
        df_colab = pd.read_excel(uploaded_file, sheet_name=1)
        df_cat = pd.read_excel(uploaded_file, sheet_name=2)

        st.info("🔄 Procesando estructura de fechas y normalizando catálogo...")

        # --- PROCESAMIENTO DE FECHAS (Hoja 1) ---
        # Las fechas están en la fila índice 1 (segunda fila)
        # Limpiamos valores nulos y nos quedamos con las columnas de fechas
        row_fechas = df_raw_turnos.iloc[1, 1:]
        fechas_limpias = []
        for f in row_fechas:
            if pd.isna(f):
                fechas_limpias.append("Sin_Fecha")
            else:
                try:
                    # Intentamos convertir a fecha de Excel, si no, a string
                    fechas_limpias.append(pd.to_datetime(f).strftime('%d-%m-%Y'))
                except:
                    fechas_limpias.append(str(f).strip())

        # Reestructuramos la matriz de turnos
        df_turnos = df_raw_turnos.iloc[2:].copy()
        # Asignamos nombres de columnas: La primera es el nombre, el resto las fechas
        df_turnos.columns = ['Colaborador_Original'] + fechas_limpias
        # Eliminamos filas donde el nombre esté vacío
        df_turnos = df_turnos.dropna(subset=['Colaborador_Original'])

        # --- NORMALIZACIÓN DEL CATÁLOGO (Hoja 3) ---
        # Columna 0: Sigla, Columna 1: Horario
        df_cat['horario_norm'] = df_cat.iloc[:, 1].apply(normalizar_turno)
        # Creamos el mapa: { "08:00 - 19:00": "E-AM1" }
        mapa_turnos = dict(zip(df_cat['horario_norm'], df_cat.iloc[:, 0]))
        mapa_turnos["L"] = "L" # Caso especial para Libres

        # --- CRUCE DE COLABORADORES (Hoja 2) ---
        # Limpiamos base de colaboradores
        df_colab.columns = [str(c).strip() for c in df_colab.columns]
        nombres_base = df_colab.iloc[:, 0].astype(str).str.upper().tolist()
        
        # Diccionario para mapear Nombre Corto -> Nombre Completo
        nombres_unicos_matriz = df_turnos['Colaborador_Original'].unique()
        with st.spinner('Realizando Fuzzy Matching de nombres...'):
            dict_nombres = {n: encontrar_mejor_coincidencia(n, nombres_base) for n in nombres_unicos_matriz}
        
        df_turnos['Nombre Completo'] = df_turnos['Colaborador_Original'].map(dict_nombres)

        # --- TRADUCCIÓN DE TURNOS A SIGLAS ---
        # Solo procesamos las columnas que realmente son fechas
        cols_fechas_reales = [c for c in fechas_limpias if c != "Sin_Fecha"]
        
        for col in cols_fechas_reales:
            # Primero normalizamos el texto de la celda, luego mapeamos a la sigla
            df_turnos[col] = df_turnos[col].astype(str).apply(normalizar_turno)
            df_turnos[col] = df_turnos[col].map(mapa_turnos).fillna(df_turnos[col])

        # --- MERGE FINAL ---
        # Tomamos datos de identificación del colaborador
        info_colab = df_colab.iloc[:, [0, 1, 2, 3]].copy()
        info_colab.columns = ['Nombre Completo', 'RUT', 'Área', 'Supervisor']
        
        # MUY IMPORTANTE: Eliminar duplicados en la base para evitar que Alexis (u otros) se dupliquen
        info_colab = info_colab.drop_duplicates(subset=['RUT'])
        
        # Unimos la información base con los turnos procesados
        resultado = pd.merge(
            info_colab, 
            df_turnos[['Nombre Completo'] + cols_fechas_reales], 
            on='Nombre Completo', 
            how='inner'
        )

        # --- INTERFAZ Y DESCARGA ---
        st.success(f"✅ Se han vinculado {len(resultado)} colaboradores correctamente.")
        
        # Mostrar vista previa
        st.subheader("Vista Previa del Importador")
        st.dataframe(resultado.head(15), use_container_width=True)

        # Crear el archivo Excel en memoria
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            resultado.to_excel(writer, index=False, sheet_name='Importador_BUK')
            # Formatear un poco el excel (opcional)
            workbook = writer.book
            worksheet = writer.sheets['Importador_BUK']
            header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
            for col_num, value in enumerate(resultado.columns.values):
                worksheet.write(0, col_num, value, header_format)

        st.download_button(
            label="📥 Descargar Excel para BUK",
            data=output.getvalue(),
            file_name="Importador_Turnos_BUK.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error crítico en el procesamiento: {e}")
        st.info("Asegúrate de que el archivo tenga las 3 hojas en el orden correcto.")

else:
    st.info("Sube el archivo Excel para comenzar el diagnóstico y procesamiento.")
