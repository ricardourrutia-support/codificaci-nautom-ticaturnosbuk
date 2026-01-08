import streamlit as st
import pandas as pd
from unidecode import unidecode
import io
import re # Importamos Regex para trabajar con patrones numéricos

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador de Turnos BUK", layout="wide")

st.title("✈️ Transformador de Turnos para BUK (Versión Mejorada)")
st.markdown("""
Esta herramienta normaliza los horarios detectando horas y minutos numéricamente, 
ignorando si dice 'Diurno', si tiene segundos extra o diferencias de ceros (8:00 vs 08:00).
""")

# --- FUNCIONES DE LIMPIEZA ---

def limpiar_texto(texto):
    """Normaliza nombres: mayúsculas, sin acentos, sin espacios extra."""
    if not isinstance(texto, str):
        return str(texto)
    return unidecode(texto.upper().strip())

def normalizar_horario_estricto(texto_horario):
    """
    Extrae las horas y minutos usando inteligencia de patrones (Regex).
    Convierte cualquier formato (8:00, 08:00:00, 8:00 PM) a un estándar 'HH:MM-HH:MM'.
    """
    # 1. Manejo de nulos o Libre
    if pd.isna(texto_horario):
        return "LIBRE"
    
    s = str(texto_horario).upper().strip()
    if "LIBRE" in s:
        return "LIBRE"
    
    # 2. Buscar patrones de hora: "un digito o dos", seguido de ":", seguido de "dos digitos"
    # Esto captura 8:00, 08:00, 20:00. Ignora los segundos (:00) si vienen después.
    patron = r"(\d{1,2}):(\d{2})"
    coincidencias = re.findall(patron, s)
    
    # Esperamos encontrar al menos 2 horas (entrada y salida)
    if len(coincidencias) >= 2:
        # Tomamos la primera coincidencia (Entrada) y la última (Salida)
        h1, m1 = coincidencias[0]
        h2, m2 = coincidencias[1] # Ojo: Si hay colación entre medio, esto toma la segunda hora encontrada
        
        # Formateamos rellenando con ceros: 8 pasa a 08
        hora_inicio = f"{int(h1):02d}:{m1}"
        hora_fin = f"{int(h2):02d}:{m2}"
        
        return f"{hora_inicio}-{hora_fin}"
    
    return "ERROR_FORMATO"

def buscar_rut(nombre_corto, df_colaboradores):
    """
    Busca el RUT en la base de colaboradores usando coincidencia parcial del nombre.
    """
    if pd.isna(nombre_corto): return "SIN NOMBRE"
    
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
        # Intento de desempate: si el apellido exacto está
        return "REVISAR: Múltiples nombres similares"
    else:
        return "REVISAR: Nombre no encontrado"

# --- INTERFAZ DE CARGA ---

uploaded_file = st.file_uploader("Sube el archivo Excel (Turnos, Base y Codificación)", type=["xlsx"])

if uploaded_file:
    try:
        # Leer las hojas
        xls = pd.ExcelFile(uploaded_file)
        
        # Cargar DataFrames
        df_turnos = pd.read_excel(xls, sheet_name='Turnos Formato Supervisor', header=1) 
        df_colab = pd.read_excel(xls, sheet_name='Base de Colaboradores')
        df_codigos = pd.read_excel(xls, sheet_name='Codificación de Turnos')

        st.success("Archivo cargado. Normalizando horarios...")

        # --- PREPROCESAMIENTO ---

        # 1. Preparar Base de Colaboradores
        df_colab['Nombre Clean'] = df_colab['Nombre del Colaborador'].apply(limpiar_texto)
        
        # 2. Preparar Diccionario de Turnos INTELIGENTE
        # Convertimos la columna de horarios del maestro al mismo formato estandarizado
        diccionario_turnos = {}
        
        # Creamos una lista para mostrar qué códigos se cargaron (para debug)
        codigos_cargados = []

        for _, row in df_codigos.iterrows():
            horario_original = row['Horario']
            sigla = row['Sigla']
            
            # Aplicamos la misma normalización estricta al maestro
            clave_estandar = normalizar_horario_estricto(horario_original)
            
            if clave_estandar != "ERROR_FORMATO":
                diccionario_turnos[clave_estandar] = sigla
                codigos_cargados.append(f"{clave_estandar} -> {sigla}")
            
            # Caso especial: Agregar "LIBRE" manualmente si no viene
            diccionario_turnos["LIBRE"] = "L" 

        with st.expander("Ver mapeo de horarios detectados (Debug)"):
            st.write(codigos_cargados)

        # --- PROCESAMIENTO PRINCIPAL ---
        
        datos_finales = []
        
        # Columnas de fechas (saltando la primera que es nombres)
        cols_fechas = df_turnos.columns[1:] 
        
        for idx, row in df_turnos.iterrows():
            nombre_supervisor = row.iloc[0]
            if pd.isna(nombre_supervisor): continue
            
            # Buscar RUT
            rut = buscar_rut(nombre_supervisor, df_colab)
            
            fila_buk = {'RUT': rut, 'Nombre': nombre_supervisor}
            
            # Procesar cada día
            for col_fecha in cols_fechas:
                valor_celda = row[col_fecha]
                
                # 1. Normalizar lo que escribió el supervisor
                horario_normalizado = normalizar_horario_estricto(valor_celda)
                
                # 2. Buscar en el diccionario
                if horario_normalizado in diccionario_turnos:
                    sigla = diccionario_turnos[horario_normalizado]
                else:
                    # Si falla, devolvemos el horario normalizado para que veas por qué falló
                    sigla = f"NO_EXISTE: {horario_normalizado}"
                
                # Formatear nombre columna fecha
                nombre_columna_fecha = str(col_fecha).split()[0]
                fila_buk[nombre_columna_fecha] = sigla
            
            datos_finales.append(fila_buk)

        # Crear DataFrame Final
        df_final = pd.DataFrame(datos_finales)

        # Ordenar columnas
        cols = ['RUT', 'Nombre'] + [c for c in df_final.columns if c not in ['RUT', 'Nombre']]
        df_final = df_final[cols]

        # --- MOSTRAR RESULTADOS Y ERRORES ---
        
        st.subheader("Vista Previa")
        st.dataframe(df_final.head())
        
        # Detección de errores visual
        celdas_error = df_final.apply(lambda x: x.astype(str).str.contains("NO_EXISTE").any(), axis=1)
        errores_turnos = df_final[celdas_error]
        
        celdas_rut = df_final[df_final['RUT'].astype(str).str.contains("REVISAR")]

        if not errores_turnos.empty:
            st.error(f"❌ Aún hay {len(errores_turnos)} filas con turnos no reconocidos.")
            st.write("Mira la tabla de abajo para ver qué horarios estandarizados no están en tu maestro de códigos:")
            st.dataframe(errores_turnos)
        else:
            st.success("✅ ¡Todos los turnos fueron reconocidos exitosamente!")

        if not celdas_rut.empty:
             st.warning(f"⚠️ Hay problemas con {len(celdas_rut)} RUTs.")
             st.dataframe(celdas_rut[['Nombre', 'RUT']])

        # --- BOTÓN DE DESCARGA ---
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Carga BUK')
            
        st.download_button(
            label="📥 Descargar Excel Listo",
            data=buffer.getvalue(),
            file_name="Carga_Masiva_BUK_V2.xlsx",
            mime="application/vnd.ms-excel"
        )

    except Exception as e:
        st.error(f"Error crítico: {e}")

else:
    st.info("Sube tu archivo Excel actualizado.")
