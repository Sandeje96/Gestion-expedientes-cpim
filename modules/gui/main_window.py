import tkinter as tk
import customtkinter as ctk
from modules.data_manager import DataManager
from modules.file_manager import FileManager
from modules.word_generator import WordGenerator
from config import TRABAJOS_PATH, ensure_directories
from .new_record_window import NewRecordWindow
from .edit_work_window import EditWorkWindow
from .duplicate_work_window import DuplicateWorkWindow
from .generate_word_window import GenerateWordWindow


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configuración inicial
        self.title("Sistema de Gestión CPIM")
        self.geometry("1200x700")
        self.minsize(900, 600)
        
        # Inicializar managers
        self.data_manager = DataManager()
        self.file_manager = FileManager()
        self.word_generator = WordGenerator()

        # Asegurarse de que las carpetas existan
        ensure_directories()
        
        # Variables para las ventanas
        self.current_window = None
        
        # Crear menú principal
        self.create_main_menu()
    
    def create_main_menu(self):
        """Crea el menú principal de la aplicación"""
        # Limpiar la ventana
        self.clear_window()
        
        # Crear frame principal
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        title = ctk.CTkLabel(
            main_frame, 
            text="Sistema de Gestión - Consejo Profesional de Ingeniería de Misiones", 
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title.pack(pady=20)
        
        # Frame para los botones
        btn_frame = ctk.CTkFrame(main_frame)
        btn_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Configurar el grid para que sea responsive
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)
        btn_frame.grid_rowconfigure(0, weight=1)
        btn_frame.grid_rowconfigure(1, weight=1)
        btn_frame.grid_rowconfigure(2, weight=1)
        
        # Botón para nuevo registro
        btn_new = ctk.CTkButton(
            btn_frame, 
            text="Nuevo Registro", 
            font=ctk.CTkFont(size=16),
            height=80,
            command=self.show_new_record_window
        )
        btn_new.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        # Botón para editar trabajo
        btn_edit = ctk.CTkButton(
            btn_frame, 
            text="Editar Trabajo", 
            font=ctk.CTkFont(size=16),
            height=80,
            command=self.show_edit_work_window
        )
        btn_edit.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        
        # Botón para duplicar trabajo
        btn_duplicate = ctk.CTkButton(
            btn_frame, 
            text="Repetir Trabajo\ncon Otro Profesional", 
            font=ctk.CTkFont(size=16),
            height=80,
            command=self.show_duplicate_work_window
        )
        btn_duplicate.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")
        
        # Botón para generar documento Word
        btn_word = ctk.CTkButton(
            btn_frame, 
            text="Generar Documento Word", 
            font=ctk.CTkFont(size=16),
            height=80,
            command=self.show_generate_word_window
        )
        btn_word.grid(row=1, column=1, padx=20, pady=20, sticky="nsew")
        
        # Botón para ver carpeta de archivos
        btn_folder = ctk.CTkButton(
            btn_frame, 
            text="Ver Carpeta de Archivos", 
            font=ctk.CTkFont(size=16),
            height=80,
            command=lambda: self.file_manager.open_folder(TRABAJOS_PATH)
        )
        btn_folder.grid(row=2, column=0, columnspan=2, padx=20, pady=20, sticky="nsew")
    
    def clear_window(self):
        """Limpia todos los widgets de la ventana"""
        for widget in self.winfo_children():
            widget.destroy()
        self.current_window = None
    
    def show_new_record_window(self):
        """Muestra la ventana para registrar un nuevo trabajo"""
        self.clear_window()
        self.current_window = NewRecordWindow(
            self, 
            self.data_manager, 
            self.file_manager, 
            self.create_main_menu
        )
    
    def show_edit_work_window(self):
        """Muestra la ventana para editar un trabajo existente"""
        self.clear_window()
        self.current_window = EditWorkWindow(
            self, 
            self.data_manager, 
            self.file_manager, 
            self.create_main_menu
        )
    
    def show_duplicate_work_window(self):
        """Muestra la ventana para duplicar un trabajo"""
        self.clear_window()
        self.current_window = DuplicateWorkWindow(
            self, 
            self.data_manager, 
            self.file_manager, 
            self.create_main_menu,
            self.show_new_record_window
        )
    
    def show_generate_word_window(self):
        """Muestra la ventana para generar documentos Word"""
        self.clear_window()
        self.current_window = GenerateWordWindow(
            self, 
            self.data_manager, 
            self.file_manager, 
            self.word_generator,
            self.create_main_menu
        )