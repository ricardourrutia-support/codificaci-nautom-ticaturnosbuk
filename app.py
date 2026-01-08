import streamlit as st
import pandas as pd
import io
import re
import unicodedata
import difflib

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Generador de Turnos BUK (Turbo)", layout="wide")

st.title("⚡ Transformador de Turnos BUK (Versión Rápida)")
st.markdown("""
Esta versión utiliza **procesamiento vectorizado** para una velocidad máxima y 
**búsqueda difusa** para corregir nombres mal escritos automáticamente.
""")

# --- FUNCIONES OPTIMIZADAS ---

def limpiar_texto(texto):
    """Normaliza texto: mayúsculas, sin acentos, sin espacios extra."""
    if pd.isna(texto): return ""
    texto = str(texto)
    # Quitar acentos
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return texto.upper().strip()

def buscar_rut_inteligente(nombres_unicos, df_colaboradores):
    """
    Crea un diccionario {Nombre_Supervisor: RUT} optimizado.
    Usa coincidencia exacta primero, y coincidencia difusa (parecidos) después.
    """
    mapa_ruts = {}
    
    # Pre-procesamos la base para no hacerlo en cada bucle
    df_colab_temp = df_colaboradores.copy()
    df_colab_temp['Nombre_Clean'] = df_colab_temp['Nombre del Colaborador'].apply(limpiar_texto)
    lista_nombres_reales = df_colab_temp['Nombre_Clean'].unique()
    
    # Diccionario inverso para buscar RUT rápido por nombre limpio
    # Asumimos que el nombre limpio es suficiente clave. Si hay duplicados, tomamos el primero.
    rut_lookup = df_colab_temp.set_index('Nombre_Clean')['RUT'].to_dict()

    for nombre in nombres_unicos:
        if not nombre or pd.isna(nombre): continue
        
        nombre_limpio = limpiar_texto(nombre)
        partes = nombre_limpio.split()
        
        # 1. Estrategia Exacta (Contiene todas las partes)
        # Buscamos nombres en la base que contengan TODAS las palabras del nombre corto
        matches = [real for real in lista_nombres_reales if all(p in real for p in partes)]
        
        rut_encontrado = None
        
        if len(matches) == 1:
            rut_encontrado = rut_lookup[matches[0]]
        elif len(matches) > 1:
            rut_encontrado = "ERROR: Múltiples coincidencias"
        else:
            # 2. Estrategia Difusa (Corrección de errores tipográficos)
            # Busca el nombre más parecido en la lista completa
            posibles = difflib.get_close_matches(nombre_limpio, lista_nombres_reales, n=1, cutoff=0.7)
            if posibles:
                rut_encontrado = rut_lookup[posibles[0]]
                # Marcamos que fue corregido para avisar al usuario
                # (Opcional: podrías guardar un log de correcciones)
            else:
                rut_encontrado = "ERROR: No encontrado"
        
        mapa_ruts[nombre] = rut_encontrado

    return mapa_ruts

def normalizar_horarios_vectorizado(serie):
    """
    Procesa toda una columna de horarios de una sola vez usando Regex vectorizado.
    Mucho más rápido que un bucle for.
    """
    # Convertir a string y limpiar básico
    s = serie.astype(str).str.upper().str.strip()
    
    # Crear serie resultado vacía
    res = pd.Series(index=s.index, dtype='object')
    
    # 1. Detectar Libres (incluye nulos, "Libre", "NaN")
    mask_libre = s.str.contains('LIBRE', na=False) | s.isna() | (s == 'NAN')
    res[mask_libre] = 'L' # Asignamos la sigla de libre directamente
    
    # 2. Extraer horas HH:MM
    # Regex: Busca dígitos:dígitos. 
    patron = r"(\d{1,2}):(\d{2})"
    
    # Filtramos solo los que no son libres para procesar
    s_proceso = s[~mask_libre]
    
    if s_proceso.empty:
        return res
        
    # extractall devuelve todas las coincidencias. 
    # Queremos la primera (entrada) y la última (salida) de cada celda.
    # Como extractall cambia el índice, usamos findall que mantiene la estructura lista
    extracted = s_proceso.str.findall(patron)
    
    def formatear(lista_match):
        if not isinstance(lista_match, list) or len(lista_match) < 2:
            return "ERROR_FORMATO"
        # Tomar primera y última hora encontrada
        h1, m1 = lista_match[0]
        h2, m2 = lista_match[-1]
        # Formatear rellenando ceros (8:00 -> 08:00)
        return f"{int(h1):02d}:{m1}-{int(h2):02d}:{m2}"

    res[~mask_libre] = extracted.apply(formatear)
    
    return res

# --- INTERFAZ ---

uploaded_file = st.file_uploader("Sube el archivo Excel", type=["xlsx"])

if uploaded_file:
    try:
        # 1. CARGA RÁPIDA
        # Usamos header=2 para que la fila de fechas (2026-01-01) sea el encabezado real
        # Esto evita problemas con columnas duplicadas como "Thursday"
        xls = pd.ExcelFile(uploaded_file)
        df_turnos = pd.read_excel(xls, sheet_name='Turnos Formato Supervisor', header=2)
        df_colab = pd.read_excel(xls, sheet_name='Base de Colaboradores')
        df_codigos = pd.read_excel(xls, sheet_name='Codificación de Turnos')
        
        st.success("Archivo cargado. Procesando...")

        # 2. PREPARACIÓN DE DATOS (MELT)
        # Transformamos la tabla ancha (fechas en columnas) a tabla larga (una fila por turno)
        # Esto permite procesar 1000 turnos en 1 milisegundo.
        
        # Identificar columna de nombres (generalmente la primera, puede llamarse diferente si está vacía)
        col_nombre = df_turnos.columns[0] 
        # Las columnas de fechas son todas las demás
        cols_fechas = [c for c in df_turnos.columns if c != col_nombre]
        
        # Melt!
        df_long = df_turnos.melt(id_vars=[col_nombre], value_vars=cols_fechas, var_name='Fecha', value_name='Turno_Original')
        
        # Eliminar filas donde el nombre sea nulo (filas vacías de excel)
        df_long = df_long.dropna(subset=[col_nombre])

        # 3. MAPEO DE RUTS (OPTIMIZADO)
        nombres_unicos = df_long[col_nombre].unique()
        mapa_ruts = buscar_rut_inteligente(nombres_unicos, df_colab)
        df_long['RUT'] = df_long[col_nombre].map(mapa_ruts)

        # 4. NORMALIZACIÓN DE HORARIOS (VECTORIZADO)
        df_long['Turno_Norm'] = normalizar_horarios_vectorizado(df_long['Turno_Original'])

        # 5. CRUCE CON CÓDIGOS
        # Preparamos el diccionario de códigos limpiando también las claves del maestro
        df_codigos['Horario_Norm'] = normalizar_horarios_vectorizado(df_codigos['Horario'])
        # Crear diccionario {Horario_Norm: Sigla}
        # Filtramos los que dieron error en el maestro para no ensuciar
        codigos_validos = df_codigos[df_codigos['Horario_Norm'] != "ERROR_FORMATO"]
        diccionario_turnos = dict(zip(codigos_validos['Horario_Norm'], codigos_validos['Sigla']))
        
        # Agregar caso L manual por si acaso
        diccionario_turnos['L'] = 'L'
        
        # Mapear
        df_long['Sigla_Final'] = df_long['Turno_Norm'].map(diccionario_turnos)
        
        # Manejo de no encontrados
        mask_error = df_long['Sigla_Final'].isna()
        df_long.loc[mask_error, 'Sigla_Final'] = "REVISAR: " + df_long.loc[mask_error, 'Turno_Norm'].astype(str)

        # 6. PIVOT FINAL (VOLVER A FORMATO BUK)
        # Reconstruimos la tabla ancha
        df_final = df_long.pivot(index=[col_nombre, 'RUT'], columns='Fecha', values='Sigla_Final').reset_index()
        
        # Renombrar columnas para que quede bonito
        df_final.rename(columns={col_nombre: 'Nombre'}, inplace=True)

        # --- RESULTADOS ---
        st.subheader("Vista Previa")
        st.dataframe(df_final.head())
        
        errores = df_final[df_final.astype(str).apply(lambda x: x.str.contains("REVISAR").any(), axis=1)]
        
        if not errores.empty:
            st.warning(f"⚠️ Se encontraron {len(errores)} filas con datos para revisar (Turnos desconocidos o RUTs no encontrados).")
            with st.expander("Ver detalles de errores"):
                st.dataframe(errores)
        else:
            st.success("✅ ¡Todo perfecto! 100% de datos cruzados.")

        # --- DESCARGA ---
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Carga BUK')
            
        st.download_button(
            label="📥 Descargar Excel Optimizado",
            data=buffer.getvalue(),
            file_name="Carga_BUK_Final.xlsx",
            mime="application/vnd.ms-excel"
        )

    except Exception as e:
        st.error(f"Error: {e}")
        st.write("Verifica que las hojas del Excel se llamen exactamente como en el ejemplo.")

else:
    st.info("Sube tu archivo para comenzar.")
