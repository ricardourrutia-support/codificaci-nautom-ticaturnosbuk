import streamlit as st
import pandas as pd
from unidecode import unidecode
import io

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador de Turnos BUK", layout="wide")

st.title("✈️ Transformador de Turnos para BUK")
st.markdown("""
Esta herramienta toma el **Formato Supervisor** y lo convierte en el formato de carga masiva para **BUK**.
Realiza el cruce de nombres por RUT y transforma los rangos horarios en Siglas.
""")

# --- FUNCIONES DE LIMPIEZA ---

def limpiar_texto(texto):
    """Normaliza texto: mayúsculas, sin acentos, sin espacios extra."""
    if not isinstance(texto, str):
        return str(texto)
    return unidecode(texto.upper().strip())

def limpiar_horario(horario_str):
    """
    Toma un string de horario sucio (ej: '9:00:00 - 20:00:00 Diurno')
    y lo convierte a un formato estándar para comparación (ej: '09:00:00-20:00:00').
    """
    if pd.isna(horario_str) or horario_str == "Libre":
        return "LIBRE"
    
    # Quitar palabras como "Diurno", "Nocturno" y espacios
    s = limpiar_texto(horario_str)
    s = s.replace("DIURNO", "").replace("NOCTURNO", "").replace(" ", "")
    
    # Normalizar separador
    # A veces viene como 09:00-20:00 y a veces 09:00:00-20:00:00
    # Vamos a intentar estandarizar quitando segundos si existen para comparar HH:MM
    # O simplemente dejándolo limpio de letras.
    
    return s

def buscar_rut(nombre_corto, df_colaboradores):
    """
    Busca el RUT en la base de colaboradores usando coincidencia parcial del nombre.
    Ej: 'Genesis Olivero' busca coincidencia en 'GENESIS VICTORIA OLIVERO MELEAN'
    """
    nombre_clean = limpiar_texto(nombre_corto)
    partes = nombre_clean.split()
    
    # Filtrar posibles candidatos
    candidatos = df_colaboradores.copy()
    candidatos['Coincidencia'] = candidatos['Nombre Clean'].apply(
        lambda x: all(parte in x for parte in partes)
    )
    
    resultado = candidatos[candidatos['Coincidencia']]
    
    if len(resultado) == 1:
        return resultado.iloc[0]['RUT']
    elif len(resultado) > 1:
        return "ERROR: Múltiples coincidencias"
    else:
        return "ERROR: No encontrado"

def obtener_sigla(horario_sucio, dict_turnos):
    """
    Busca la sigla correspondiente al horario sucio.
    """
    # 1. Caso directo Libre
    if pd.isna(horario_sucio): return "L"
    if str(horario_sucio).strip().upper() == "LIBRE": return "L"

    # 2. Limpieza para comparación
    h_clean = limpiar_horario(horario_sucio)
    
    # Intentar buscar en el diccionario
    # El diccionario tiene claves limpias también
    if h_clean in dict_turnos:
        return dict_turnos[h_clean]
    
    # Si falla, intentar variaciones (ej: agregar o quitar segundos)
    # Esta parte es crítica si los formatos no son idénticos
    return "REVISAR" # Retorna esto si no encuentra el turno

# --- INTERFAZ DE CARGA ---

uploaded_file = st.file_uploader("Sube el archivo Excel (Turnos Formato Supervisor)", type=["xlsx"])

if uploaded_file:
    try:
        # Leer las hojas
        xls = pd.ExcelFile(uploaded_file)
        
        # Cargar DataFrames
        df_turnos = pd.read_excel(xls, sheet_name='Turnos Formato Supervisor', header=1) # Header 1 porque la fila 0 son meses
        df_colab = pd.read_excel(xls, sheet_name='Base de Colaboradores')
        df_codigos = pd.read_excel(xls, sheet_name='Codificación de Turnos')

        st.success("Archivo cargado correctamente. Procesando datos...")

        # --- PREPROCESAMIENTO ---

        # 1. Preparar Base de Colaboradores
        df_colab['Nombre Clean'] = df_colab['Nombre del Colaborador'].apply(limpiar_texto)
        
        # 2. Preparar Diccionario de Turnos (Horario -> Sigla)
        # Limpiamos la clave (Horario) igual que limpiaremos los datos de entrada
        diccionario_turnos = {}
        for _, row in df_codigos.iterrows():
            clave = limpiar_horario(row['Horario'])
            valor = row['Sigla']
            diccionario_turnos[clave] = valor
            
            # Hack: A veces Excel se come los segundos o los pone.
            # Agregamos variaciones si es necesario.
            # Por ahora confiamos en la función limpiar_horario

        # --- PROCESAMIENTO PRINCIPAL ---
        
        datos_finales = []
        
        # Iterar sobre la matriz de turnos
        # La columna 0 es el nombre, el resto son fechas
        cols_fechas = df_turnos.columns[1:] 
        
        for idx, row in df_turnos.iterrows():
            nombre_supervisor = row.iloc[0] # Primera columna
            if pd.isna(nombre_supervisor): continue
            
            # Buscar RUT
            rut = buscar_rut(nombre_supervisor, df_colab)
            
            fila_buk = {'RUT': rut, 'Nombre': nombre_supervisor}
            
            # Procesar cada día
            for col_fecha in cols_fechas:
                valor_celda = row[col_fecha]
                sigla = obtener_sigla(valor_celda, diccionario_turnos)
                
                # Formatear la fecha como string si es datetime
                nombre_columna_fecha = str(col_fecha).split()[0] # Tomar solo YYYY-MM-DD
                fila_buk[nombre_columna_fecha] = sigla
            
            datos_finales.append(fila_buk)

        # Crear DataFrame Final
        df_final = pd.DataFrame(datos_finales)

        # Mover RUT al principio
        cols = ['RUT', 'Nombre'] + [c for c in df_final.columns if c not in ['RUT', 'Nombre']]
        df_final = df_final[cols]

        # --- MOSTRAR RESULTADOS Y ERRORES ---
        
        st.subheader("Vista Previa del Archivo para BUK")
        st.dataframe(df_final.head())
        
        # Detección de errores (Turnos no encontrados o RUTs no encontrados)
        errores_rut = df_final[df_final['RUT'].str.contains("ERROR", na=False)]
        
        # Verificar si hay siglas "REVISAR"
        celdas_revisar = df_final.apply(lambda x: x.astype(str).str.contains("REVISAR").any(), axis=1)
        errores_turnos = df_final[celdas_revisar]

        if not errores_rut.empty:
            st.warning(f"⚠️ Hay {len(errores_rut)} colaboradores cuyo RUT no se encontró.")
            st.dataframe(errores_rut[['Nombre', 'RUT']])
            
        if not errores_turnos.empty:
            st.warning(f"⚠️ Hay {len(errores_turnos)} filas con turnos que no coincidieron con el maestro.")
            st.write("Busca las celdas que dicen 'REVISAR'. Probablemente el horario escrito no existe en la hoja 'Codificación de Turnos'.")

        # --- BOTÓN DE DESCARGA ---
        
        # Convertir a Excel en memoria
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Carga BUK')
            
        st.download_button(
            label="📥 Descargar Excel procesado para BUK",
            data=buffer.getvalue(),
            file_name="Carga_Masiva_BUK.xlsx",
            mime="application/vnd.ms-excel"
        )

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
        st.write("Por favor verifica que el archivo tenga las 3 hojas con los nombres correctos.")

else:
    st.info("Por favor sube el archivo Excel para comenzar.")
