import streamlit as st
import pandas as pd
import io
import utils  # Importamos el módulo local utils.py

# Configuración de la página
st.set_page_config(page_title="BUK Shift Importer PRO", layout="wide")

st.title("🚀 Generador de Importador BUK")
st.markdown("""
**Instrucciones:**
1. Sube el archivo Excel con las 3 hojas (Turnos, Colaboradores, Configuración).
2. El sistema normalizará nombres y horarios automáticamente.
3. Descarga el archivo resultante listo para BUK.
""")

# Carga de archivo
uploaded_file = st.file_uploader("Carga el Excel 'Turnos Formato Supervisor'", type=["xlsx"])

if uploaded_file:
    try:
        # ---------------------------------------------------------
        # 1. CARGA DE DATOS ESTRUCTURAL
        # ---------------------------------------------------------
        # Leemos la Hoja 1 sin encabezado para localizar las fechas manualmente
        df_raw_turnos = pd.read_excel(uploaded_file, sheet_name=0, header=None)
        df_colab = pd.read_excel(uploaded_file, sheet_name=1)
        df_cat = pd.read_excel(uploaded_file, sheet_name=2)

        st.info("🔄 Analizando estructura del archivo...")

        # ---------------------------------------------------------
        # 2. PROCESAMIENTO DE FECHAS (HOJA 1)
        # ---------------------------------------------------------
        # Las fechas suelen estar en la fila índice 1 (fila 2 de Excel)
        row_fechas = df_raw_turnos.iloc[1, 1:].values
        
        fechas_columnas = []
        conteo_fechas = {} # Para evitar columnas duplicadas

        for f in row_fechas:
            if pd.isna(f):
                col_name = "VACIO"
            else:
                try:
                    # Intentamos convertir a string de fecha DD-MM-YYYY
                    col_name = pd.to_datetime(f).strftime('%d-%m-%Y')
                except:
                    col_name = str(f).strip()
            
            # Manejo de duplicados (Ej: si hay dos columnas '01-01-2026')
            if col_name in conteo_fechas:
                conteo_fechas[col_name] += 1
                col_name = f"{col_name}_{conteo_fechas[col_name]}"
            else:
                conteo_fechas[col_name] = 0
            
            fechas_columnas.append(col_name)

        # Reconstruimos el DataFrame de Turnos
        df_turnos = df_raw_turnos.iloc[2:].copy()
        # Asignamos nombres: Primera col es Nombre, resto son las fechas procesadas
        df_turnos.columns = ['Nombre_Corto'] + fechas_columnas
        
        # Eliminamos filas vacías o basura
        df_turnos = df_turnos[df_turnos['Nombre_Corto'].notna()]
        
        # Filtramos solo las columnas que parecen fechas válidas (tienen guiones o slashes)
        cols_fechas_reales = [c for c in fechas_columnas if "-" in c and "VACIO" not in c]

        # ---------------------------------------------------------
        # 3. PREPARACIÓN DEL CATÁLOGO (HOJA 3)
        # ---------------------------------------------------------
        # Normalizamos la columna de horarios del catálogo
        df_cat['horario_norm'] = df_cat.iloc[:, 1].apply(utils.normalizar_turno)
        
        # Creamos diccionario: { "08:00 - 20:00": "SIGLA" }
        mapa_turnos = dict(zip(df_cat['horario_norm'], df_cat.iloc[:, 0]))
        mapa_turnos["L"] = "L" # Aseguramos que Libre se mapee a L

        # ---------------------------------------------------------
        # 4. FUZZY MATCHING DE NOMBRES (HOJA 1 VS HOJA 2)
        # ---------------------------------------------------------
        # Limpieza de nombres en base de colaboradores
        df_colab.columns = [str(c).strip() for c in df_colab.columns]
        nombres_base = df_colab.iloc[:, 0].astype(str).str.upper().tolist()
        
        nombres_unicos_matriz = df_turnos['Nombre_Corto'].unique()
        
        # Barra de progreso visual
        texto_progreso = st.empty()
        bar_progreso = st.progress(0)
        
        dict_nombres = {}
        total = len(nombres_unicos_matriz)
        
        for i, nombre in enumerate(nombres_unicos_matriz):
            # Usamos la función del utils
            match = utils.encontrar_mejor_coincidencia(nombre, nombres_base)
            dict_nombres[nombre] = match
            # Actualizar barra cada cierto tiempo
            if i % 5 == 0 or i == total:
                bar_progreso.progress((i + 1) / total)
                texto_progreso.text(f"Analizando nombre: {nombre}...")
        
        texto_progreso.empty()
        bar_progreso.empty()

        # Aplicamos el mapeo de nombres
        df_turnos['Nombre Completo'] = df_turnos['Nombre_Corto'].map(dict_nombres)

        # ---------------------------------------------------------
        # 5. TRADUCCIÓN DE TURNOS (CORE LOGIC)
        # ---------------------------------------------------------
        st.info("Traducidos horarios a códigos BUK...")
        
        for col in cols_fechas_reales:
            # 1. Convertir a string y normalizar (limpiar basura, regex horas)
            serie_normalizada = df_turnos[col].astype(str).apply(utils.normalizar_turno)
            
            # 2. Mapear contra el catálogo
            serie_mapeada = serie_normalizada.map(mapa_turnos)
            
            # 3. Rellenar los que no encontraron mapa con el valor normalizado original
            # (Esto soluciona el error "value must be scalar" usando combine_first o fillna seguro)
            df_turnos[col] = serie_mapeada.fillna(serie_normalizada)

        # ---------------------------------------------------------
        # 6. UNIÓN FINAL (MERGE)
        # ---------------------------------------------------------
        # Preparamos la base de colaboradores (Hoja 2)
        info_colab = df_colab.iloc[:, [0, 1, 2, 3]].copy()
        info_colab.columns = ['Nombre Completo', 'RUT', 'Área', 'Supervisor']
        
        # IMPORTANTE: Eliminar duplicados por RUT en la base maestra para evitar explosión de filas
        info_colab = info_colab.drop_duplicates(subset=['RUT'])
        
        # Merge: Unimos info personal con la matriz de turnos
        # Solo traemos las filas donde hubo coincidencia de nombre
        resultado = pd.merge(
            info_colab, 
            df_turnos[['Nombre Completo'] + cols_fechas_reales], 
            on='Nombre Completo', 
            how='inner'
        )

        # ---------------------------------------------------------
        # 7. EXPORTACIÓN
        # ---------------------------------------------------------
        if not resultado.empty:
            st.success(f"✅ ¡Proceso exitoso! {len(resultado)} colaboradores procesados.")
            
            st.subheader("Vista Previa")
            st.dataframe(resultado.head(10))

            # Configuración de descarga Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                resultado.to_excel(writer, index=False, sheet_name='Importador_BUK')
                
                # Ajustes de formato visual en el Excel
                workbook = writer.book
                worksheet = writer.sheets['Importador_BUK']
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
                
                for col_num, value in enumerate(resultado.columns.values):
                    worksheet.write(0, col_num, value, header_fmt)
                    worksheet.set_column(col_num, col_num, 15) # Ancho de columna

            st.download_button(
                label="📥 Descargar Excel Final (.xlsx)",
                data=output.getvalue(),
                file_name="Importador_Turnos_BUK.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ No se generaron registros. Revisa que los nombres de la Hoja 1 coincidan mínimamente con la Hoja 2.")

    except Exception as e:
        st.error("⛔ Ocurrió un error crítico durante el procesamiento.")
        st.code(f"Detalle del error: {str(e)}")
        st.markdown("Recomendación: Verifica que el archivo Excel tenga las 3 hojas en el orden correcto y que no tenga columnas con nombres duplicados.")
