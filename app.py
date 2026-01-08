import streamlit as st
import pandas as pd
import io
import re
import unicodedata
import difflib

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador de Turnos BUK (Full Datos)", layout="wide")

st.title("✈️ Transformador de Turnos BUK (Formato Completo)")
st.markdown("""
Esta herramienta procesa los turnos, corrige nombres mal escritos y **agrega automáticamente** los datos maestros (Área, Supervisor) de cada colaborador.
""")

# --- FUNCIONES DE LIMPIEZA Y LÓGICA ---

def limpiar_texto(texto):
    """Normaliza texto para comparaciones (quita acentos, mayúsculas, espacios)."""
    if pd.isna(texto): return ""
    texto = str(texto)
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return texto.upper().strip()

def buscar_rut_inteligente(nombres_unicos, df_colaboradores):
    """
    Busca el RUT basándose en el nombre corto del supervisor.
    Usa coincidencia exacta primero, y coincidencia difusa (parecidos) después.
    """
    mapa_ruts = {}
    
    # Pre-procesamos la base
    df_colab_temp = df_colaboradores.copy()
    df_colab_temp['Nombre_Clean'] = df_colab_temp['Nombre del Colaborador'].apply(limpiar_texto)
    lista_nombres_reales = df_colab_temp['Nombre_Clean'].unique()
    
    # Diccionario para buscar RUT rápido
    rut_lookup = df_colab_temp.set_index('Nombre_Clean')['RUT'].to_dict()

    for nombre in nombres_unicos:
        if not nombre or pd.isna(nombre): continue
        
        nombre_limpio = limpiar_texto(nombre)
        partes = nombre_limpio.split()
        
        # 1. Estrategia Exacta (Todas las palabras coinciden)
        matches = [real for real in lista_nombres_reales if all(p in real for p in partes)]
        
        rut_encontrado = None
        
        if len(matches) == 1:
            rut_encontrado = rut_lookup[matches[0]]
        elif len(matches) > 1:
            rut_encontrado = "ERROR: Múltiples coincidencias"
        else:
            # 2. Estrategia Difusa (Corrección de errores tipográficos)
            posibles = difflib.get_close_matches(nombre_limpio, lista_nombres_reales, n=1, cutoff=0.7)
            if posibles:
                rut_encontrado = rut_lookup[posibles[0]]
            else:
                rut_encontrado = "ERROR: No encontrado"
        
        mapa_ruts[nombre] = rut_encontrado

    return mapa_ruts

def normalizar_horarios_vectorizado(serie):
    """Convierte cualquier formato de hora a HH:MM-HH:MM de forma masiva."""
    s = serie.astype(str).str.upper().str.strip()
    res = pd.Series(index=s.index, dtype='object')
    
    # Detectar Libres
    mask_libre = s.str.contains('LIBRE', na=False) | s.isna() | (s == 'NAN')
    res[mask_libre] = 'L'
    
    # Regex para HH:MM
    patron = r"(\d{1,2}):(\d{2})"
    s_proceso = s[~mask_libre]
    
    if s_proceso.empty:
        return res
        
    extracted = s_proceso.str.findall(patron)
    
    def formatear(lista_match):
        if not isinstance(lista_match, list) or len(lista_match) < 2:
            return "ERROR_FORMATO"
        h1, m1 = lista_match[0]
        h2, m2 = lista_match[-1]
        return f"{int(h1):02d}:{m1}-{int(h2):02d}:{m2}"

    res[~mask_libre] = extracted.apply(formatear)
    return res

# --- INTERFAZ PRINCIPAL ---

uploaded_file = st.file_uploader("Sube el archivo Excel (con las 3 hojas)", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # 1. CARGA DE DATOS
        # Hoja 1: Turnos (usamos header=2 para tomar la fila de fechas)
        df_turnos = pd.read_excel(xls, sheet_name='Turnos Formato Supervisor', header=2)
        # Hoja 2: Base Maestra
        df_colab = pd.read_excel(xls, sheet_name='Base de Colaboradores')
        # Hoja 3: Códigos
        df_codigos = pd.read_excel(xls, sheet_name='Codificación de Turnos')
        
        st.success("Archivo cargado. Procesando y cruzando bases de datos...")

        # 2. TRANSFORMACIÓN INICIAL (MELT)
        col_nombre_input = df_turnos.columns[0]
        cols_fechas = [c for c in df_turnos.columns if c != col_nombre_input]
        
        df_long = df_turnos.melt(id_vars=[col_nombre_input], value_vars=cols_fechas, var_name='Fecha', value_name='Turno_Original')
        df_long = df_long.dropna(subset=[col_nombre_input])

        # 3. OBTENCIÓN DE RUT (INTELIGENTE)
        nombres_unicos = df_long[col_nombre_input].unique()
        mapa_ruts = buscar_rut_inteligente(nombres_unicos, df_colab)
        df_long['RUT'] = df_long[col_nombre_input].map(mapa_ruts)

        # 4. CRUCE CON BASE DE COLABORADORES (VLOOKUP)
        # Aquí traemos Área, Supervisor y el Nombre Completo Oficial
        # Aseguramos que las columnas existan en df_colab
        cols_maestras = ['Nombre del Colaborador', 'RUT', 'Área', 'Supervisor']
        # Filtramos solo lo que nos sirve de la base
        df_base_clean = df_colab[cols_maestras].copy()
        
        # Hacemos el Merge (unir) usando el RUT como llave
        df_merged = pd.merge(df_long, df_base_clean, on='RUT', how='left')

        # Si el nombre oficial no se encuentra (porque el RUT falló), usamos el nombre corto del input para no perder la fila
        df_merged['Nombre del Colaborador'] = df_merged['Nombre del Colaborador'].fillna(df_merged[col_nombre_input])

        # 5. NORMALIZACIÓN DE HORARIOS Y CÓDIGOS
        df_merged['Turno_Norm'] = normalizar_horarios_vectorizado(df_merged['Turno_Original'])

        # Preparar diccionario códigos
        df_codigos['Horario_Norm'] = normalizar_horarios_vectorizado(df_codigos['Horario'])
        codigos_validos = df_codigos[df_codigos['Horario_Norm'] != "ERROR_FORMATO"]
        diccionario_turnos = dict(zip(codigos_validos['Horario_Norm'], codigos_validos['Sigla']))
        diccionario_turnos['L'] = 'L'
        
        df_merged['Sigla_Final'] = df_merged['Turno_Norm'].map(diccionario_turnos)
        
        # Marcar errores
        mask_error = df_merged['Sigla_Final'].isna()
        df_merged.loc[mask_error, 'Sigla_Final'] = "REVISAR: " + df_merged.loc[mask_error, 'Turno_Norm'].astype(str)

        # 6. PIVOT FINAL (ESTRUCTURA BUK COMPLETA)
        # Usamos las columnas maestras como índice para que queden a la izquierda
        cols_index = ['Nombre del Colaborador', 'RUT', 'Área', 'Supervisor']
        
        # Rellenar vacíos en Área/Supervisor por estética si falló el cruce
        df_merged['Área'] = df_merged['Área'].fillna("Desconocido")
        df_merged['Supervisor'] = df_merged['Supervisor'].fillna("Desconocido")

        df_final = df_merged.pivot(index=cols_index, columns='Fecha', values='Sigla_Final').reset_index()

        # --- MOSTRAR RESULTADOS ---
        st.subheader("Vista Previa (Formato BUK Final)")
        st.dataframe(df_final.head())
        
        # Control de Calidad
        errores_rut = df_final[df_final['RUT'].astype(str).str.contains("ERROR")]
        if not errores_rut.empty:
            st.warning(f"⚠️ Atención: {len(errores_rut)} colaboradores no cruzaron con la base maestra.")
            st.write("Estos aparecerán con Área 'Desconocido'.")

        # --- DESCARGA ---
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Carga BUK')
            
            # Ajuste de ancho de columnas (Opcional, para que se vea bonito al abrir)
            workbook  = writer.book
            worksheet = writer.sheets['Carga BUK']
            format_left = workbook.add_format({'align': 'left'})
            worksheet.set_column('A:A', 30, format_left) # Nombre
            worksheet.set_column('B:B', 12, format_left) # RUT
            worksheet.set_column('C:D', 20, format_left) # Área y Supervisor
            
        st.download_button(
            label="📥 Descargar Excel Listo para BUK",
            data=buffer.getvalue(),
            file_name="Reporte_Turnos_BUK_Completo.xlsx",
            mime="application/vnd.ms-excel"
        )

    except Exception as e:
        st.error(f"Ocurrió un error: {e}")
        st.write("Detalle técnico: Verifica que la hoja 'Base de Colaboradores' tenga las columnas: 'Nombre del Colaborador', 'RUT', 'Área', 'Supervisor'.")

else:
    st.info("Sube el archivo Excel para procesar.")
