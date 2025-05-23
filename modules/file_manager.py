import os
import shutil
import subprocess
from pathlib import Path
from config import TRABAJOS_PATH, INFORMES_PATH

class FileManager:
    def __init__(self):
        self.trabajos_path = TRABAJOS_PATH
    
    def create_folder_structure(self, data, tipo_trabajo):
        """
        Crea la estructura de carpetas para un trabajo
        
        Args:
            data: Diccionario con datos del trabajo
            tipo_trabajo: 'obra' o 'informe'
        
        Returns:
            Path: Ruta a la carpeta creada
        """
        if tipo_trabajo == "obra":
            # Determinar la carpeta base dependiendo del tipo de trabajo
            if data["tipo_trabajo"] == "Obra nueva":
                base_folder = self.trabajos_path / "OBRA NUEVA"
            else:  # Registración
                base_folder = self.trabajos_path / "REGISTRACION"
            
            # Crear carpeta del profesional
            profesional_folder = base_folder / data["nombre_profesional"]
            profesional_folder.mkdir(exist_ok=True)
            
            # Crear carpeta del comitente
            comitente_folder = profesional_folder / data["nombre_comitente"]
            comitente_folder.mkdir(exist_ok=True)
            
            return comitente_folder
        
        elif tipo_trabajo == "informe":
            # Para informes técnicos, usamos la nueva estructura
            from config import INFORMES_PATH
            
            # Crear un nombre seguro para la carpeta del tipo de informe
            tipo_informe = data["tipo_trabajo"]
            safe_name = tipo_informe.replace("/", "-").replace("\\", "-").replace(":", " -")
            
            # Determinar la carpeta del tipo de informe
            tipo_informe_folder = INFORMES_PATH / safe_name
            tipo_informe_folder.mkdir(exist_ok=True)
            
            # Crear carpeta del profesional
            profesional_folder = tipo_informe_folder / data["profesional"]
            profesional_folder.mkdir(exist_ok=True)
            
            # Crear carpeta del comitente
            comitente_folder = profesional_folder / data["comitente"]
            comitente_folder.mkdir(exist_ok=True)
            
            return comitente_folder
    
    def open_folder(self, folder_path):
        """Abre la carpeta en el explorador de archivos"""
        if os.path.exists(folder_path):
            # Usar el comando específico del sistema operativo
            if os.name == 'nt':  # Windows
                os.startfile(folder_path)
            elif os.name == 'posix':  # macOS y Linux
                if 'darwin' in os.sys.platform:  # macOS
                    subprocess.call(['open', folder_path])
                else:  # Linux
                    subprocess.call(['xdg-open', folder_path])
            return True
        return False