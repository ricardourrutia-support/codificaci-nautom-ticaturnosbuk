import re
import pandas as pd

def normalizar_turno(texto):
    """Limpia y estandariza horarios de forma segura."""
    # Validación robusta para valores nulos o vacíos
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ""
    
    # Convertir a string de forma segura
    texto_str = str(texto).strip().upper()
    
    if texto_str in ["NAN", "", "NONE"]:
        return ""
    
    # 1. Quitar etiquetas Diurno/Nocturno
    texto_str = texto_str.replace("DIURNO", "").replace("NOCTURNO", "").strip()
    
    # 2. Quitar segundos (:00)
    texto_str = re.sub(r":00(?!\d)", "", texto_str)
    
    # 3. Forzar formato HH:MM (8:00 -> 08:00)
    texto_str = re.sub(r"(\d{1,2}):(\d{2})", lambda m: f"{int(m.group(1)):02d}:{m.group(2)}", texto_str)
    
    # 4. Estandarizar el separador guion
    texto_str = re.sub(r"\s*[-–]\s*", " - ", texto_str)
    
    # 5. Manejo de Libres
    if "LIBRE" in texto_str or texto_str == "L":
        return "L"
        
    return texto_str
