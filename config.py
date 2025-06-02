import os
import sys
from pathlib import Path

# Función para obtener la ruta base del proyecto
def get_base_path():
    # Si se ejecuta como script congelado (exe)
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    # Si se ejecuta como script .py
    return Path(__file__).parent

# Rutas principales
BASE_PATH = get_base_path()
DATA_PATH = BASE_PATH / "data"
TEMPLATES_PATH = BASE_PATH / "templates"
TRABAJOS_PATH = BASE_PATH / "trabajos"
INFORMES_PATH = BASE_PATH / "informes_tecnicos"  # Nueva carpeta

# Rutas específicas
EXCEL_FILE = DATA_PATH / "registros.xlsx"

TIPOS_INFORME = [
    "Informe de Homologación-Cambio de tipo", 
    "Plan de Contingencia", 
    "Informe de Caldera", 
    "Informe de Trailer-Casa Rodante",
    "Informe Técnico"
]

# Rutas para plantillas Word
TEMPLATE_OBRA_SELLADO = TEMPLATES_PATH / "obra_general_sellado.docx"
TEMPLATE_OBRA_VISADO = TEMPLATES_PATH / "obra_general_visado.docx"
TEMPLATE_INFORME = TEMPLATES_PATH / "informe_tecnico.docx"

# Crear las carpetas si no existen
def ensure_directories():
    # Crear carpetas principales
    for path in [DATA_PATH, TEMPLATES_PATH, TRABAJOS_PATH, INFORMES_PATH]:
        try:
            path.mkdir(exist_ok=True)
            print(f"Carpeta creada o verificada: {path}")
        except Exception as e:
            print(f"Error al crear carpeta {path}: {e}")
    
    # Crear carpetas específicas para trabajos
    try:
        (TRABAJOS_PATH / "OBRA NUEVA").mkdir(exist_ok=True)
        (TRABAJOS_PATH / "REGISTRACION").mkdir(exist_ok=True)
        print("Estructura de carpetas para trabajos creada correctamente")
    except Exception as e:
        print(f"Error al crear estructura de carpetas para trabajos: {e}")
    
    # Crear carpetas para cada tipo de informe técnico
    try:
        for tipo_informe in TIPOS_INFORME:
            # Crear un nombre seguro para carpetas
            safe_name = tipo_informe.replace("/", "-").replace("\\", "-").replace(":", " -")
            (INFORMES_PATH / safe_name).mkdir(exist_ok=True)
        print("Estructura de carpetas para informes técnicos creada correctamente")
    except Exception as e:
        print(f"Error al crear estructura de carpetas para informes técnicos: {e}")


# Constantes
TIPOS_OBRA = ["Obra nueva", "Registración"]
PROFESIONES = ["Ingeniero", "Licenciado", "MMO", "Técnico"]
FORMATOS = ["Físico", "Digital"]
WHATSAPP_WEB_URL = "https://web.whatsapp.com/send"