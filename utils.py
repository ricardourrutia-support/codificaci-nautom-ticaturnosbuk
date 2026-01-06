import re
import pandas as pd
from fuzzywuzzy import process

def normalizar_turno(texto):
    """Limpia y estandariza horarios."""
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ""
    
    texto_str = str(texto).strip().upper()
    if texto_str in ["NAN", "", "NONE", "LIBRE"]:
        return "L"
    
    # Estandarización de formato
    texto_str = texto_str.replace("DIURNO", "").replace("NOCTURNO", "").strip()
    texto_str = re.sub(r":00(?!\d)", "", texto_str)
    texto_str = re.sub(r"(\d{1,2}):(\d{2})", lambda m: f"{int(m.group(1)):02d}:{m.group(2)}", texto_str)
    texto_str = re.sub(r"\s*[-–]\s*", " - ", texto_str)
        
    return texto_str

def encontrar_mejor_coincidencia(nombre_corto, lista_nombres_completos):
    """Busca coincidencia de nombres."""
    if not nombre_corto or pd.isna(nombre_corto):
        return None
    
    nombre_corto = str(nombre_corto).strip().upper()
    # Umbral de 80% para evitar errores
    resultado = process.extractOne(nombre_corto, lista_nombres_completos)
    if resultado and resultado[1] >= 80:
        return resultado[0]
    return None
