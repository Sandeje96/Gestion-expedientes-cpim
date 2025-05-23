import os
import shutil
import subprocess
import re
from pathlib import Path
from config import TRABAJOS_PATH, INFORMES_PATH

class FileManager:
    def __init__(self):
        self.trabajos_path = TRABAJOS_PATH
    
    def sanitize_folder_name(self, name):
        """
        Limpia el nombre de una carpeta para que sea válido en el sistema de archivos
        
        Args:
            name: Nombre original de la carpeta
        
        Returns:
            str: Nombre limpio y válido para carpeta
        """
        if not name:
            return "Sin_Nombre"
        
        # Limpiar espacios al inicio y final
        clean_name = name.strip()
        
        # Si después de limpiar está vacío, usar un nombre por defecto
        if not clean_name:
            return "Sin_Nombre"
        
        # Reemplazar caracteres problemáticos para nombres de carpeta
        # Caracteres no válidos en Windows: < > : " | ? * \ /
        # También evitamos puntos al final y espacios múltiples
        invalid_chars = r'[<>:"|?*\\\/]'
        clean_name = re.sub(invalid_chars, '_', clean_name)
        
        # Reemplazar múltiples espacios por uno solo
        clean_name = re.sub(r'\s+', ' ', clean_name)
        
        # Reemplazar espacios por guiones bajos (opcional, puedes cambiarlo)
        # clean_name = clean_name.replace(' ', '_')
        
        # Eliminar puntos al final (problemático en Windows)
        clean_name = clean_name.rstrip('.')
        
        # Limitar longitud del nombre (opcional)
        if len(clean_name) > 100:
            clean_name = clean_name[:100]
        
        # Si después de todo esto está vacío, usar nombre por defecto
        if not clean_name:
            return "Sin_Nombre"
        
        return clean_name
    
    def create_folder_structure(self, data, tipo_trabajo):
        """
        Crea la estructura de carpetas para un trabajo
        
        Args:
            data: Diccionario con datos del trabajo
            tipo_trabajo: 'obra' o 'informe'
        
        Returns:
            Path: Ruta a la carpeta creada
        """
        try:
            if tipo_trabajo == "obra":
                # Determinar la carpeta base dependiendo del tipo de trabajo
                if data["tipo_trabajo"] == "Obra nueva":
                    base_folder = self.trabajos_path / "OBRA NUEVA"
                else:  # Registración
                    base_folder = self.trabajos_path / "REGISTRACION"
                
                # Limpiar y crear carpeta del profesional
                profesional_name = self.sanitize_folder_name(data["nombre_profesional"])
                profesional_folder = base_folder / profesional_name
                profesional_folder.mkdir(parents=True, exist_ok=True)
                
                # Limpiar y crear carpeta del comitente
                comitente_name = self.sanitize_folder_name(data["nombre_comitente"])
                comitente_folder = profesional_folder / comitente_name
                comitente_folder.mkdir(parents=True, exist_ok=True)
                
                return comitente_folder
            
            elif tipo_trabajo == "informe":
                # Para informes técnicos, usamos la nueva estructura
                
                # Crear un nombre seguro para la carpeta del tipo de informe
                tipo_informe = data["tipo_trabajo"]
                safe_name = self.sanitize_folder_name(tipo_informe)
                
                # Determinar la carpeta del tipo de informe
                tipo_informe_folder = INFORMES_PATH / safe_name
                tipo_informe_folder.mkdir(parents=True, exist_ok=True)
                
                # Limpiar y crear carpeta del profesional
                profesional_name = self.sanitize_folder_name(data["profesional"])
                profesional_folder = tipo_informe_folder / profesional_name
                profesional_folder.mkdir(parents=True, exist_ok=True)
                
                # Limpiar y crear carpeta del comitente
                comitente_name = self.sanitize_folder_name(data["comitente"])
                comitente_folder = profesional_folder / comitente_name
                comitente_folder.mkdir(parents=True, exist_ok=True)
                
                return comitente_folder
        
        except Exception as e:
            print(f"Error al crear estructura de carpetas: {e}")
            # En caso de error, crear una carpeta de fallback
            fallback_folder = self.trabajos_path / "ERROR_CARPETAS" / f"{tipo_trabajo}_{data.get('fecha', 'sin_fecha')}"
            fallback_folder.mkdir(parents=True, exist_ok=True)
            return fallback_folder
    
    def open_folder(self, folder_path):
        """Abre la carpeta en el explorador de archivos"""
        try:
            folder_path = Path(folder_path)
            
            # Verificar que la carpeta existe
            if not folder_path.exists():
                print(f"La carpeta no existe: {folder_path}")
                return False
            
            # Usar el comando específico del sistema operativo
            if os.name == 'nt':  # Windows
                os.startfile(str(folder_path))
            elif os.name == 'posix':  # macOS y Linux
                if 'darwin' in os.sys.platform:  # macOS
                    subprocess.call(['open', str(folder_path)])
                else:  # Linux
                    subprocess.call(['xdg-open', str(folder_path)])
            
            return True
            
        except Exception as e:
            print(f"Error al abrir carpeta {folder_path}: {e}")
            return False