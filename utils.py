import re
import pandas as pd
from fuzzywuzzy import process

def normalizar_turno(texto):
    """Estandariza formatos de hora (8:00 -> 08:00) y espacios."""
    if pd.isna(texto) or str(texto).strip() == "": 
        return ""
    texto = str(texto).strip().upper()
    # 8:00 o 8:00 -> 08:00
    texto = re.sub(r"(\d{1,2}):(\d{2})", lambda m: f"{int(m.group(1)):02d}:{m.group(2)}", texto)
    # Espacios uniformes en el guion
    texto = re.sub(r"\s*-\s*", " - ", texto)
    return texto

def encontrar_mejor_coincidencia(nombre_corto, lista_nombres_completos):
    """Busca el nombre completo más parecido al nombre corto de la matriz."""
    if pd.isna(nombre_corto): return None
    resultado = process.extractOne(str(nombre_corto).upper(), lista_nombres_completos)
    if resultado and resultado[1] > 70: # Umbral de confianza del 70%
        return resultado[0]
    return None

def formatear_rut(rut):
    """Limpia el RUT para que sea consistente."""
    if pd.isna(rut): return ""
    return str(rut).strip().upper()
