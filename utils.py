import re
import pandas as pd
from fuzzywuzzy import process

def normalizar_turno(texto):
    """Limpia y estandariza horarios: '9:00:00 - 20:00:00 Diurno' -> '09:00 - 20:00'"""
    if pd.isna(texto) or str(texto).strip().lower() in ["nan", ""]: 
        return ""
    
    texto = str(texto).strip().upper()
    
    # 1. Quitar etiquetas Diurno/Nocturno
    texto = texto.replace("DIURNO", "").replace("NOCTURNO", "").strip()
    
    # 2. Quitar segundos (:00) si existen
    texto = re.sub(r":00(?!\d)", "", texto)
    
    # 3. Forzar formato HH:MM (8:00 -> 08:00)
    texto = re.sub(r"(\d{1,2}):(\d{2})", lambda m: f"{int(m.group(1)):02d}:{m.group(2)}", texto)
    
    # 4. Estandarizar el separador guion
    texto = re.sub(r"\s*[-–]\s*", " - ", texto)
    
    # 5. Manejo especial de Libres
    if "LIBRE" in texto or texto == "L":
        return "L"
        
    return texto

def encontrar_mejor_coincidencia(nombre_corto, lista_nombres_completos):
    if pd.isna(nombre_corto): return None
    # Buscamos con un umbral alto para evitar duplicados falsos
    resultado = process.extractOne(str(nombre_corto).upper(), lista_nombres_completos)
    if resultado and resultado[1] > 80:
        return resultado[0]
    return None
