import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
from datetime import datetime
from pathlib import Path
import subprocess
import shutil
from tkinter import filedialog
from tkinter import TclError
from modules.data_manager import DataManager
from modules.file_manager import FileManager
from modules.word_generator import WordGenerator
from config import PROFESIONES, FORMATOS, TIPOS_OBRA, TIPOS_INFORME, TRABAJOS_PATH, ensure_directories


class AutocompleteEntry(ctk.CTkFrame):
    """Campo de entrada con autocompletado"""
    def __init__(self, parent, options=None, width=200, height=30, **kwargs):
        super().__init__(parent, width=width, height=height)
        
        self.parent = parent
        self.options = options or []
        self.matches = []
        self.entry_var = ctk.StringVar()
        self.dropdown_open = False
        
        # Crear el widget de entrada
        self.entry = ctk.CTkEntry(self, textvariable=self.entry_var, width=width, height=height, **kwargs)
        self.entry.pack(fill="x", expand=True)
        
        # Crear una ventana de nivel superior para el dropdown
        self.toplevel = None
        self.listbox = None
        
        # Configurar eventos
        self.entry_var.trace_add("write", self.on_entry_change)
        self.entry.bind("<FocusOut>", self.on_focus_out)
        self.entry.bind("<FocusIn>", self.on_focus_in)
        
        # Vincular teclas de navegación
        self.entry.bind("<Down>", self.on_down)
        self.entry.bind("<Up>", self.on_up)
        self.entry.bind("<Return>", self.on_return)
        self.entry.bind("<Escape>", lambda e: self.hide_dropdown())
    
    def create_dropdown(self):
        """Crea la ventana del dropdown"""
        if self.toplevel is not None:
            self.toplevel.destroy()
        
        # Crear ventana de nivel superior sin decoraciones
        self.toplevel = tk.Toplevel(self)
        self.toplevel.overrideredirect(True)  # Sin bordes ni título
        self.toplevel.attributes('-topmost', True)  # Siempre al frente
        
        # Aplicar temas de color
        self.toplevel.configure(bg="#F0F0F0")
        
        # Crear el listbox
        self.listbox = tk.Listbox(self.toplevel, font=("Arial", 12), bg="#F0F0F0",
                                  highlightthickness=1, highlightbackground="#CCCCCC")
        self.listbox.pack(fill="both", expand=True)
        
        # Vincular la selección
        self.listbox.bind("<<ListboxSelect>>", self.on_select)
        
        # Vincular teclas dentro del listbox
        self.listbox.bind("<Return>", self.on_return)
        self.listbox.bind("<Escape>", lambda e: self.hide_dropdown())
        
    def on_entry_change(self, *args):
        """Maneja el cambio en el texto del campo de entrada"""
        search_text = self.entry_var.get().lower()
        if not search_text:
            self.hide_dropdown()
            return
        
        # Filtrar opciones que coinciden
        self.matches = [opt for opt in self.options if search_text in opt.lower()]
        
        # Si hay coincidencias, mostrar el dropdown
        if self.matches and self.winfo_ismapped():
            # Crear dropdown si no existe
            if self.toplevel is None or not self.toplevel.winfo_exists():
                self.create_dropdown()
            
            # Limpiar y rellenar la lista
            self.listbox.delete(0, tk.END)
            for match in self.matches:
                self.listbox.insert(tk.END, match)
            
            # Limitar la altura del dropdown según número de elementos
            items_count = min(len(self.matches), 5)
            self.listbox.config(height=items_count)
            
            # Posicionar el dropdown debajo del campo de entrada
            x, y, width, height = self.get_dropdown_position()
            self.toplevel.geometry(f"{width}x{height}+{x}+{y}")
            
            # Asegurar que sea visible
            self.toplevel.deiconify()
            self.dropdown_open = True
        else:
            self.hide_dropdown()
    
    def get_dropdown_position(self):
        """Calcula la posición absoluta del dropdown"""
        # Coordenadas absolutas del widget en pantalla
        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height()
        width = self.entry.winfo_width()
        height = min(len(self.matches), 5) * 24  # Aproximadamente 24 píxeles por elemento
        
        return x, y, width, height
    
    def on_focus_out(self, event):
        """Oculta el dropdown cuando se pierde el foco"""
        # Damos un pequeño retraso para permitir la selección en la lista
        self.after(100, self.hide_dropdown)
    
    def hide_dropdown(self):
        """Oculta el dropdown"""
        if self.toplevel is not None and self.toplevel.winfo_exists():
            self.toplevel.withdraw()
        self.dropdown_open = False
    
    def on_focus_in(self, event):
        """Muestra el dropdown cuando se obtiene el foco si hay texto"""
        if self.entry_var.get():
            self.on_entry_change()
    
    def on_select(self, event):
        """Maneja la selección de un elemento del dropdown"""
        if self.listbox.curselection():
            index = self.listbox.curselection()[0]
            self.entry_var.set(self.matches[index])
            self.hide_dropdown()
            self.entry.focus_set()
    
    def on_down(self, event):
        """Maneja la tecla de flecha abajo"""
        if not self.dropdown_open:
            # Si el dropdown no está abierto, mostrarlo
            self.on_entry_change()
            if self.dropdown_open and self.listbox:
                # Seleccionar el primer elemento
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(0)
                self.listbox.activate(0)
                self.listbox.see(0)
        elif self.listbox:
            # Si está abierto, moverse al siguiente elemento
            try:
                current = self.listbox.curselection()[0]
                next_idx = current + 1
                if next_idx < len(self.matches):
                    self.listbox.selection_clear(0, tk.END)
                    self.listbox.selection_set(next_idx)
                    self.listbox.activate(next_idx)
                    self.listbox.see(next_idx)
            except (IndexError, TclError):
                # Si no hay selección, seleccionar el primer elemento
                if len(self.matches) > 0:
                    self.listbox.selection_set(0)
                    self.listbox.activate(0)
                    self.listbox.see(0)
        return "break"  # Evita que el evento se propague
    
    def on_up(self, event):
        """Maneja la tecla de flecha arriba"""
        if self.dropdown_open and self.listbox:
            try:
                current = self.listbox.curselection()[0]
                prev_idx = current - 1
                if prev_idx >= 0:
                    self.listbox.selection_clear(0, tk.END)
                    self.listbox.selection_set(prev_idx)
                    self.listbox.activate(prev_idx)
                    self.listbox.see(prev_idx)
            except (IndexError, TclError):
                # Si no hay selección, seleccionar el último elemento
                if len(self.matches) > 0:
                    last_idx = len(self.matches) - 1
                    self.listbox.selection_set(last_idx)
                    self.listbox.activate(last_idx)
                    self.listbox.see(last_idx)
        return "break"  # Evita que el evento se propague
    
    def on_return(self, event):
        """Maneja la tecla Enter para seleccionar un elemento"""
        if self.dropdown_open and self.listbox:
            try:
                # Si hay un elemento seleccionado
                if self.listbox.curselection():
                    index = self.listbox.curselection()[0]
                    self.entry_var.set(self.matches[index])
                    self.hide_dropdown()
                    self.entry.focus_set()
                    return "break"  # Evita que el evento se propague
            except (IndexError, TclError):
                pass
        return None  # Permite que el evento se propague normalmente
    
    def get(self):
        """Retorna el valor actual del campo"""
        return self.entry_var.get()
    
    def set(self, value):
        """Establece el valor del campo"""
        self.entry_var.set(value)
    
    def update_options(self, options):
        """Actualiza la lista de opciones para autocompletado"""
        self.options = options or []

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
        
        # Crear frames principales
        self.create_main_menu()
    
    def create_main_menu(self):
        """Crea el menú principal de la aplicación"""
        # Limpiar la ventana
        for widget in self.winfo_children():
            widget.destroy()
        
        # Crear frame principal
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        title = ctk.CTkLabel(main_frame, text="Sistema de Gestión - Consejo Profesional de Ingeniería de Misiones", 
                            font=ctk.CTkFont(size=24, weight="bold"))
        title.pack(pady=20)
        
        # Botones
        btn_frame = ctk.CTkFrame(main_frame)
        btn_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Botón para nuevo registro
        btn_new = ctk.CTkButton(btn_frame, text="Nuevo Registro", 
                               font=ctk.CTkFont(size=16),
                               height=80,
                               command=self.show_new_record_window)
        btn_new.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        # Botón para editar trabajo
        btn_edit = ctk.CTkButton(btn_frame, text="Editar Trabajo", 
                                font=ctk.CTkFont(size=16),
                                height=80,
                                command=self.show_edit_work_window)
        btn_edit.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        
        # Botón para generar documento Word
        btn_word = ctk.CTkButton(btn_frame, text="Generar Documento Word", 
                                font=ctk.CTkFont(size=16),
                                height=80,
                                command=self.show_generate_word_window)
        btn_word.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")
        
        # Botón para ver carpeta de archivos
        btn_folder = ctk.CTkButton(btn_frame, text="Ver Carpeta de Archivos", 
                                  font=ctk.CTkFont(size=16),
                                  height=80,
                                  command=lambda: self.file_manager.open_folder(TRABAJOS_PATH))
        btn_folder.grid(row=1, column=1, padx=20, pady=20, sticky="nsew")
        
        # Configurar el grid para que sea responsive
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)
        btn_frame.grid_rowconfigure(0, weight=1)
        btn_frame.grid_rowconfigure(1, weight=1)
    
    def show_new_record_window(self):
        """Muestra la ventana para registrar un nuevo trabajo"""
        # Limpiar la ventana
        for widget in self.winfo_children():
            widget.destroy()
        
        # Crear frame principal
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        title = ctk.CTkLabel(main_frame, text="Nuevo Registro", 
                            font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=10)
        
        # Crear pestañas para los diferentes tipos de trabajo
        tab_view = ctk.CTkTabview(main_frame)
        tab_view.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Añadir pestañas
        tab_obra = tab_view.add("Obra en general")
        tab_informe = tab_view.add("Informe técnico")
        
        # Contenido de la pestaña "Obra en general"
        self.setup_obra_tab(tab_obra)
        
        # Contenido de la pestaña "Informe técnico"
        self.setup_informe_tab(tab_informe)
        
        # Botón para volver al menú principal
        btn_back = ctk.CTkButton(main_frame, text="Volver al Menú Principal", 
                                font=ctk.CTkFont(size=14),
                                command=self.create_main_menu)
        btn_back.pack(pady=10)
    
    def setup_obra_tab(self, parent):
        """Configura los widgets para la pestaña de Obra en general"""
        # Crear frame de scroll
        scroll_frame = ctk.CTkScrollableFrame(parent)
        scroll_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Obtener listas de profesionales y comitentes existentes
        profesionales = self.data_manager.get_all_profesionales()
        comitentes = self.data_manager.get_all_comitentes()
        
        # Crear variables
        self.obra_vars = {
        "fecha": ctk.StringVar(value=datetime.now().strftime("%d/%m/%Y")),
        "profesion": ctk.StringVar(),
        "formato": ctk.StringVar(),
        "nro_copias": ctk.StringVar(),
        "tipo_trabajo": ctk.StringVar(),
        "nombre_profesional": None,  # Se usará AutocompleteEntry directamente
        "nombre_comitente": None,    # Se usará AutocompleteEntry directamente
        "ubicacion": ctk.StringVar(),
        "nro_expte_municipal": ctk.StringVar(),
        "nro_sistema_gop": ctk.StringVar(),
        "nro_partida_inmobiliaria": ctk.StringVar(),
        "tasa_sellado": ctk.StringVar(),
        "tasa_visado": ctk.StringVar(),
        "visado_gas": ctk.StringVar(),
        "visado_salubridad": ctk.StringVar(),
        "visado_electrica": ctk.StringVar(),
        "visado_electromecanica": ctk.StringVar()
    }
        
        # Crear entradas de datos
        row = 0
        
        # Fecha
        ctk.CTkLabel(scroll_frame, text="Fecha:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkEntry(scroll_frame, textvariable=self.obra_vars["fecha"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Profesión
        ctk.CTkLabel(scroll_frame, text="Profesión:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkOptionMenu(scroll_frame, values=PROFESIONES, variable=self.obra_vars["profesion"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Formato
        ctk.CTkLabel(scroll_frame, text="Formato:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        formato_frame = ctk.CTkFrame(scroll_frame)
        formato_frame.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        
        # Radio buttons para formato
        self.formato_var = ctk.StringVar(value="Físico")
        self.obra_vars["formato"] = self.formato_var
        
        ctk.CTkRadioButton(formato_frame, text="Físico", variable=self.formato_var, value="Físico", 
                        command=self.toggle_formato_fields).pack(side=tk.LEFT, padx=10)
        ctk.CTkRadioButton(formato_frame, text="Digital", variable=self.formato_var, value="Digital", 
                        command=self.toggle_formato_fields).pack(side=tk.LEFT, padx=10)
        row += 1
        
        # Número de copias (solo visible si es formato físico)
        self.nro_copias_label = ctk.CTkLabel(scroll_frame, text="Nro. de Copias:")
        self.nro_copias_label.grid(row=row, column=0, sticky="w", pady=5, padx=5)
        self.nro_copias_entry = ctk.CTkEntry(scroll_frame, textvariable=self.obra_vars["nro_copias"])
        self.nro_copias_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Número de sistema GOP (solo visible si es formato digital)
        self.nro_sistema_gop_label = ctk.CTkLabel(scroll_frame, text="Nro. de Sistema GOP:")
        self.nro_sistema_gop_label.grid(row=row, column=0, sticky="w", pady=5, padx=5)
        self.nro_sistema_gop_entry = ctk.CTkEntry(scroll_frame, textvariable=self.obra_vars["nro_sistema_gop"])
        self.nro_sistema_gop_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Frame para subir archivos (solo visible si es formato digital)
        self.upload_files_frame = ctk.CTkFrame(scroll_frame)
        self.upload_files_frame.grid(row=row, column=0, columnspan=2, sticky="ew", pady=5, padx=5)
        
        # Etiqueta para archivos
        ctk.CTkLabel(self.upload_files_frame, text="Archivos seleccionados:").pack(anchor="w", padx=5, pady=5)
        
        # Lista de archivos seleccionados
        self.files_listbox = tk.Listbox(self.upload_files_frame, height=5, width=50)
        self.files_listbox.pack(fill="x", padx=5, pady=5)
        
        # Botones para archivos
        files_buttons_frame = ctk.CTkFrame(self.upload_files_frame)
        files_buttons_frame.pack(fill="x", padx=5, pady=5)
        
        # Botón para añadir archivos
        add_files_btn = ctk.CTkButton(files_buttons_frame, text="Añadir Archivos", 
                                    command=self.add_files_to_list)
        add_files_btn.pack(side="left", padx=5)
        
        # Botón para eliminar archivos seleccionados
        remove_files_btn = ctk.CTkButton(files_buttons_frame, text="Eliminar Seleccionado", 
                                        command=self.remove_selected_file)
        remove_files_btn.pack(side="left", padx=5)
        
        row += 1
        
        # Guardar lista de archivos seleccionados
        self.selected_files = []
        
        # Ajustar visibilidad inicial de campos según formato
        self.toggle_formato_fields()
        
        # Tipo de trabajo
        ctk.CTkLabel(scroll_frame, text="Tipo de trabajo:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkOptionMenu(scroll_frame, values=TIPOS_OBRA, variable=self.obra_vars["tipo_trabajo"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Nombre del profesional (con autocompletado)
        ctk.CTkLabel(scroll_frame, text="Nombre del Profesional:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        self.prof_entry = AutocompleteEntry(scroll_frame, options=profesionales)
        self.prof_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        self.obra_vars["nombre_profesional"] = self.prof_entry
        row += 1
        
        # Nombre del comitente (con autocompletado)
        ctk.CTkLabel(scroll_frame, text="Nombre del Comitente:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        self.comit_entry = AutocompleteEntry(scroll_frame, options=comitentes)
        self.comit_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        self.obra_vars["nombre_comitente"] = self.comit_entry
        row += 1
        
        # Ubicación
        ctk.CTkLabel(scroll_frame, text="Ubicación:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkEntry(scroll_frame, textvariable=self.obra_vars["ubicacion"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Número de expediente municipal
        ctk.CTkLabel(scroll_frame, text="Nro. de Expte. Municipal:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkEntry(scroll_frame, textvariable=self.obra_vars["nro_expte_municipal"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Número de partida inmobiliaria
        ctk.CTkLabel(scroll_frame, text="Nro. de Partida Inmobiliaria:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkEntry(scroll_frame, textvariable=self.obra_vars["nro_partida_inmobiliaria"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Tasa de sellado
        ctk.CTkLabel(scroll_frame, text="Tasa de Sellado:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkEntry(scroll_frame, textvariable=self.obra_vars["tasa_sellado"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Tasa de visado (mantenemos este campo para compatibilidad)
        ctk.CTkLabel(scroll_frame, text="Tasa de Visado (General):").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkEntry(scroll_frame, textvariable=self.obra_vars["tasa_visado"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Añadir un separador o título para la sección de visados específicos
        ctk.CTkLabel(scroll_frame, text="Visados Específicos", font=("Arial", 12, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", pady=(15, 5), padx=5)
        row += 1
        
        # Visado de instalación de Gas
        ctk.CTkLabel(scroll_frame, text="Visado de instalación de Gas:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkEntry(scroll_frame, textvariable=self.obra_vars["visado_gas"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Visado de instalación de Salubridad
        ctk.CTkLabel(scroll_frame, text="Visado de instalación de Salubridad:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkEntry(scroll_frame, textvariable=self.obra_vars["visado_salubridad"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Visado de instalación eléctrica
        ctk.CTkLabel(scroll_frame, text="Visado de instalación eléctrica:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkEntry(scroll_frame, textvariable=self.obra_vars["visado_electrica"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Visado de instalación electromecánica
        ctk.CTkLabel(scroll_frame, text="Visado de instalación electromecánica:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkEntry(scroll_frame, textvariable=self.obra_vars["visado_electromecanica"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Configurar el grid para que sea responsive
        scroll_frame.grid_columnconfigure(1, weight=1)
        
        # Botón para guardar
        btn_save = ctk.CTkButton(scroll_frame, text="Guardar Obra", 
                                font=ctk.CTkFont(size=14),
                                command=self.save_obra)
        btn_save.grid(row=row, column=0, columnspan=2, pady=20, padx=5, sticky="ew")
    
    def toggle_formato_fields(self):
        """Alterna la visibilidad de los campos según el formato seleccionado"""
        formato = self.formato_var.get()
        
        if formato == "Físico":
            self.nro_copias_label.grid()
            self.nro_copias_entry.grid()
            self.nro_sistema_gop_label.grid_remove()
            self.nro_sistema_gop_entry.grid_remove()
            self.upload_files_frame.grid_remove()
        else:  # Digital
            self.nro_copias_label.grid_remove()
            self.nro_copias_entry.grid_remove()
            self.nro_sistema_gop_label.grid()
            self.nro_sistema_gop_entry.grid()
            self.upload_files_frame.grid()
            
    def add_files_to_list(self):
        """Añade archivos a la lista de archivos seleccionados"""
        filetypes = [
            ("Documentos", "*.pdf;*.doc;*.docx"),
            ("Archivos PDF", "*.pdf"),
            ("Documentos Word", "*.doc;*.docx"),
            ("Todos los archivos", "*.*")
        ]
        
        # Abrir diálogo para seleccionar archivos
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos",
            filetypes=filetypes
        )
        
        if files:
            # Añadir archivos a la lista
            for file in files:
                if file not in self.selected_files:
                    self.selected_files.append(file)
                    # Mostrar solo el nombre del archivo en la lista
                    self.files_listbox.insert(tk.END, os.path.basename(file))

    def remove_selected_file(self):
        """Elimina el archivo seleccionado de la lista"""
        selected_indices = self.files_listbox.curselection()
        
        if selected_indices:
            # Eliminar archivo seleccionado (de mayor a menor índice para evitar problemas)
            for index in sorted(selected_indices, reverse=True):
                del self.selected_files[index]
                self.files_listbox.delete(index)
    
    def setup_informe_tab(self, parent):
        """Configura los widgets para la pestaña de Informe técnico"""
        # Crear frame de scroll
        scroll_frame = ctk.CTkScrollableFrame(parent)
        scroll_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Obtener listas de profesionales y comitentes existentes
        profesionales = self.data_manager.get_all_profesionales()
        comitentes = self.data_manager.get_all_comitentes()
        
        # Crear variables
        self.informe_vars = {
            "fecha": ctk.StringVar(value=datetime.now().strftime("%d/%m/%Y")),
            "profesion": ctk.StringVar(),
            "formato": ctk.StringVar(),
            "nro_copias": ctk.StringVar(),
            "tipo_trabajo": ctk.StringVar(),
            "detalle": ctk.StringVar(),
            "profesional": None,  # Se usará AutocompleteEntry directamente
            "comitente": None,    # Se usará AutocompleteEntry directamente
            "tasa_sellado": ctk.StringVar()
        }
        
        # Crear entradas de datos
        row = 0
        
        # Fecha
        ctk.CTkLabel(scroll_frame, text="Fecha:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkEntry(scroll_frame, textvariable=self.informe_vars["fecha"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Profesión
        ctk.CTkLabel(scroll_frame, text="Profesión:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkOptionMenu(scroll_frame, values=PROFESIONES, variable=self.informe_vars["profesion"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Formato
        ctk.CTkLabel(scroll_frame, text="Formato:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        formato_frame = ctk.CTkFrame(scroll_frame)
        formato_frame.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        
        # Radio buttons para formato
        self.informe_formato_var = ctk.StringVar(value="Físico")
        self.informe_vars["formato"] = self.informe_formato_var
        
        ctk.CTkRadioButton(formato_frame, text="Físico", variable=self.informe_formato_var, value="Físico", 
                        command=self.toggle_informe_formato_fields).pack(side=tk.LEFT, padx=10)
        ctk.CTkRadioButton(formato_frame, text="Digital", variable=self.informe_formato_var, value="Digital", 
                        command=self.toggle_informe_formato_fields).pack(side=tk.LEFT, padx=10)
        row += 1
        
        # Número de copias (solo visible si es formato físico)
        self.informe_nro_copias_label = ctk.CTkLabel(scroll_frame, text="Nro. de Copias:")
        self.informe_nro_copias_label.grid(row=row, column=0, sticky="w", pady=5, padx=5)
        self.informe_nro_copias_entry = ctk.CTkEntry(scroll_frame, textvariable=self.informe_vars["nro_copias"])
        self.informe_nro_copias_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Frame para subir archivos (solo visible si es formato digital)
        self.informe_upload_files_frame = ctk.CTkFrame(scroll_frame)
        self.informe_upload_files_frame.grid(row=row, column=0, columnspan=2, sticky="ew", pady=5, padx=5)
        
        # Etiqueta para archivos
        ctk.CTkLabel(self.informe_upload_files_frame, text="Archivos seleccionados:").pack(anchor="w", padx=5, pady=5)
        
        # Lista de archivos seleccionados
        self.informe_files_listbox = tk.Listbox(self.informe_upload_files_frame, height=5, width=50)
        self.informe_files_listbox.pack(fill="x", padx=5, pady=5)
        
        # Botones para archivos
        files_buttons_frame = ctk.CTkFrame(self.informe_upload_files_frame)
        files_buttons_frame.pack(fill="x", padx=5, pady=5)
        
        # Botón para añadir archivos
        add_files_btn = ctk.CTkButton(files_buttons_frame, text="Añadir Archivos", 
                                    command=self.add_files_to_informe_list)
        add_files_btn.pack(side="left", padx=5)
        
        # Botón para eliminar archivos seleccionados
        remove_files_btn = ctk.CTkButton(files_buttons_frame, text="Eliminar Seleccionado", 
                                        command=self.remove_selected_informe_file)
        remove_files_btn.pack(side="left", padx=5)
        
        row += 1
        
        # Guardar lista de archivos seleccionados
        self.informe_selected_files = []
        
        # Ajustar visibilidad inicial de campos según formato
        self.toggle_informe_formato_fields()
        
        # Tipo de trabajo (desplegable con opciones predefinidas)
        ctk.CTkLabel(scroll_frame, text="Tipo de trabajo:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkOptionMenu(scroll_frame, values=TIPOS_INFORME, variable=self.informe_vars["tipo_trabajo"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Detalle
        ctk.CTkLabel(scroll_frame, text="Detalle:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkEntry(scroll_frame, textvariable=self.informe_vars["detalle"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Profesional (con autocompletado)
        ctk.CTkLabel(scroll_frame, text="Profesional:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        self.informe_prof_entry = AutocompleteEntry(scroll_frame, options=profesionales)
        self.informe_prof_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        self.informe_vars["profesional"] = self.informe_prof_entry
        row += 1
        
        # Comitente (con autocompletado)
        ctk.CTkLabel(scroll_frame, text="Comitente:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        self.informe_comit_entry = AutocompleteEntry(scroll_frame, options=comitentes)
        self.informe_comit_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        self.informe_vars["comitente"] = self.informe_comit_entry
        row += 1
        
        # Tasa de sellado
        ctk.CTkLabel(scroll_frame, text="Tasa de Sellado:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkEntry(scroll_frame, textvariable=self.informe_vars["tasa_sellado"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Configurar el grid para que sea responsive
        scroll_frame.grid_columnconfigure(1, weight=1)
        
        # Botón para guardar
        btn_save = ctk.CTkButton(scroll_frame, text="Guardar Informe", 
                                font=ctk.CTkFont(size=14),
                                command=self.save_informe)
        btn_save.grid(row=row, column=0, columnspan=2, pady=20, padx=5, sticky="ew")
    
    def toggle_informe_formato_fields(self):
        """Alterna la visibilidad de los campos según el formato seleccionado para informes"""
        formato = self.informe_formato_var.get()
        
        if formato == "Físico":
            self.informe_nro_copias_label.grid()
            self.informe_nro_copias_entry.grid()
            self.informe_upload_files_frame.grid_remove()
        else:  # Digital
            self.informe_nro_copias_label.grid_remove()
            self.informe_nro_copias_entry.grid_remove()
            self.informe_upload_files_frame.grid()

    def add_files_to_informe_list(self):
        """Añade archivos a la lista de archivos seleccionados para informes"""
        filetypes = [
            ("Documentos", "*.pdf;*.doc;*.docx"),
            ("Archivos PDF", "*.pdf"),
            ("Documentos Word", "*.doc;*.docx"),
            ("Todos los archivos", "*.*")
        ]
        
        # Abrir diálogo para seleccionar archivos
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos",
            filetypes=filetypes
        )
        
        if files:
            # Añadir archivos a la lista
            for file in files:
                if file not in self.informe_selected_files:
                    self.informe_selected_files.append(file)
                    # Mostrar solo el nombre del archivo en la lista
                    self.informe_files_listbox.insert(tk.END, os.path.basename(file))

    def remove_selected_informe_file(self):
        """Elimina el archivo seleccionado de la lista de informes"""
        selected_indices = self.informe_files_listbox.curselection()
        
        if selected_indices:
            # Eliminar archivo seleccionado (de mayor a menor índice para evitar problemas)
            for index in sorted(selected_indices, reverse=True):
                del self.informe_selected_files[index]
                self.informe_files_listbox.delete(index)
    
    def save_obra(self):
        """Guarda un nuevo registro de Obra en general"""
        # Validar campos obligatorios
        if not self.obra_vars["profesion"].get().strip():
            messagebox.showerror("Error", "El campo 'profesion' es obligatorio.")
            return
        if not self.obra_vars["formato"].get().strip():
            messagebox.showerror("Error", "El campo 'formato' es obligatorio.")
            return
        if not self.obra_vars["tipo_trabajo"].get().strip():
            messagebox.showerror("Error", "El campo 'tipo_trabajo' es obligatorio.")
            return
        if not self.obra_vars["nombre_profesional"].get().strip():
            messagebox.showerror("Error", "El campo 'nombre_profesional' es obligatorio.")
            return
        if not self.obra_vars["nombre_comitente"].get().strip():
            messagebox.showerror("Error", "El campo 'nombre_comitente' es obligatorio.")
            return
        
        # Obtener datos del formulario
        data = {}
        for key, var in self.obra_vars.items():
            if key in ["nombre_profesional", "nombre_comitente"]:
                # Estos son widgets AutocompleteEntry
                data[key] = var.get()
            else:
                # Estos son StringVar
                data[key] = var.get()
        
        # Crear la estructura de carpetas si es digital
        if data["formato"] == "Digital":
            folder_path = self.file_manager.create_folder_structure(data, "obra")
            data["ruta_carpeta"] = str(folder_path)
            
            # Copiar archivos seleccionados a la carpeta del comitente
            if self.selected_files:
                for file_path in self.selected_files:
                    try:
                        # Obtener solo el nombre del archivo
                        file_name = os.path.basename(file_path)
                        # Crear ruta de destino
                        destination = os.path.join(folder_path, file_name)
                        # Copiar archivo
                        shutil.copy2(file_path, destination)
                    except Exception as e:
                        messagebox.showerror("Error", f"No se pudo copiar el archivo {file_name}: {str(e)}")
        
        # Guardar en el Excel
        row_id = self.data_manager.add_obra_general(data)
        
        # Mostrar mensaje de éxito
        messagebox.showinfo("Éxito", "Obra registrada correctamente.")
        
        # Abrir la carpeta si es digital
        if data["formato"] == "Digital" and "ruta_carpeta" in data:
            self.file_manager.open_folder(data["ruta_carpeta"])
        
        # Limpiar formulario pero actualizar las opciones de autocompletado
        profesionales = self.data_manager.get_all_profesionales()
        comitentes = self.data_manager.get_all_comitentes()
        
        self.obra_vars["nombre_profesional"].update_options(profesionales)
        self.obra_vars["nombre_comitente"].update_options(comitentes)
        
        for key, var in self.obra_vars.items():
            if key != "fecha" and key not in ["nombre_profesional", "nombre_comitente"]:
                var.set("")
        
        self.obra_vars["nombre_profesional"].set("")
        self.obra_vars["nombre_comitente"].set("")
        
        # Limpiar lista de archivos
        self.selected_files = []
        self.files_listbox.delete(0, tk.END)

    def generate_informe(self, informe_id):
        """Genera un documento Word para un informe técnico"""
        # Obtener los datos del informe
        informe = self.data_manager.get_work_by_id("informe", informe_id)
        
        if informe:
            # Determinar la ruta de salida
            if informe["formato"] == "Digital" and informe["ruta_carpeta"]:
                output_dir = Path(informe["ruta_carpeta"])
            else:
                output_dir = Path.cwd()
            
            output_path = output_dir / f"Informe_{informe['comitente']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            
            # Generar el documento
            result = self.word_generator.generate_informe(informe, output_path)
            
            if result.startswith("Error") or result.startswith("Faltan"):
                messagebox.showerror("Error", result)
            else:
                messagebox.showinfo("Éxito", f"Documento generado correctamente en:\n{result}")
                
                # Abrir la carpeta donde se guardó el documento
                self.file_manager.open_folder(output_dir)
        else:
            messagebox.showerror("Error", "No se encontró la información del informe")
    
    def save_informe(self):
        """Guarda un nuevo registro de Informe técnico"""
        # Validar campos obligatorios
        if not self.informe_vars["profesion"].get().strip():
            messagebox.showerror("Error", "El campo 'profesion' es obligatorio.")
            return
        if not self.informe_vars["formato"].get().strip():
            messagebox.showerror("Error", "El campo 'formato' es obligatorio.")
            return
        if not self.informe_vars["tipo_trabajo"].get().strip():
            messagebox.showerror("Error", "El campo 'tipo_trabajo' es obligatorio.")
            return
        if not self.informe_vars["profesional"].get().strip():
            messagebox.showerror("Error", "El campo 'profesional' es obligatorio.")
            return
        if not self.informe_vars["comitente"].get().strip():
            messagebox.showerror("Error", "El campo 'comitente' es obligatorio.")
            return
        
        # Obtener datos del formulario
        data = {}
        for key, var in self.informe_vars.items():
            if key in ["profesional", "comitente"]:
                # Estos son widgets AutocompleteEntry
                data[key] = var.get()
            else:
                # Estos son StringVar
                data[key] = var.get()
        
        # Crear la estructura de carpetas si es digital
        if data["formato"] == "Digital":
            folder_path = self.file_manager.create_folder_structure(data, "informe")
            data["ruta_carpeta"] = str(folder_path)
            
            # Copiar archivos seleccionados a la carpeta del comitente
            if self.informe_selected_files:
                for file_path in self.informe_selected_files:
                    try:
                        # Obtener solo el nombre del archivo
                        file_name = os.path.basename(file_path)
                        # Crear ruta de destino
                        destination = os.path.join(folder_path, file_name)
                        # Copiar archivo
                        shutil.copy2(file_path, destination)
                    except Exception as e:
                        messagebox.showerror("Error", f"No se pudo copiar el archivo {file_name}: {str(e)}")
        
        # Guardar en el Excel
        row_id = self.data_manager.add_informe_tecnico(data)
        
        # Mostrar mensaje de éxito
        messagebox.showinfo("Éxito", "Informe registrado correctamente.")
        
        # Abrir la carpeta si es digital
        if data["formato"] == "Digital" and "ruta_carpeta" in data:
            self.file_manager.open_folder(data["ruta_carpeta"])
        
        # Limpiar formulario pero actualizar las opciones de autocompletado
        profesionales = self.data_manager.get_all_profesionales()
        comitentes = self.data_manager.get_all_comitentes()
        
        self.informe_vars["profesional"].update_options(profesionales)
        self.informe_vars["comitente"].update_options(comitentes)
        
        for key, var in self.informe_vars.items():
            if key != "fecha" and key not in ["profesional", "comitente"]:
                var.set("")
        
        self.informe_vars["profesional"].set("")
        self.informe_vars["comitente"].set("")
        
        # Limpiar lista de archivos
        self.informe_selected_files = []
        self.informe_files_listbox.delete(0, tk.END)
    
    def show_edit_work_window(self):
        """Muestra la ventana para editar un trabajo existente"""
        # Limpiar la ventana
        for widget in self.winfo_children():
            widget.destroy()
        
        # Crear frame principal
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        title = ctk.CTkLabel(main_frame, text="Editar Trabajo", 
                            font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=10)
        
        # Crear pestañas para los diferentes tipos de trabajo
        tab_view = ctk.CTkTabview(main_frame)
        tab_view.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Añadir pestañas
        tab_obra = tab_view.add("Obras en general")
        tab_informe = tab_view.add("Informes técnicos")
        
        # Configurar las pestañas
        self.setup_edit_obra_tab(tab_obra)
        self.setup_edit_informe_tab(tab_informe)
        
        # Botón para volver al menú principal
        btn_back = ctk.CTkButton(main_frame, text="Volver al Menú Principal", 
                                font=ctk.CTkFont(size=14),
                                command=self.create_main_menu)
        btn_back.pack(pady=10)
    
    def setup_edit_obra_tab(self, parent):
        """Configura la pestaña para editar obras en general"""
        # Crear frame superior para búsqueda y selección
        selection_frame = ctk.CTkFrame(parent)
        selection_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Frame para búsqueda
        search_frame = ctk.CTkFrame(selection_frame)
        search_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Etiqueta para búsqueda
        ctk.CTkLabel(search_frame, text="Buscar por:").grid(row=0, column=0, padx=5, pady=5)
        
        # Opciones de búsqueda
        search_options = ["Profesional", "Comitente", "Partida Inmobiliaria", "Número GOP"]
        self.search_option_var = ctk.StringVar(value=search_options[0])
        option_menu = ctk.CTkOptionMenu(search_frame, values=search_options, variable=self.search_option_var)
        option_menu.grid(row=0, column=1, padx=5, pady=5)
        
        # Obtener listas de profesionales y comitentes existentes
        profesionales = self.data_manager.get_all_profesionales()
        comitentes = self.data_manager.get_all_comitentes()
        
        # Campo de búsqueda con autocompletado
        self.search_entry = AutocompleteEntry(search_frame, width=200, options=profesionales)
        self.search_entry.grid(row=0, column=2, padx=5, pady=5)
        
        # Actualizar opciones de autocompletado según el criterio seleccionado
        def update_search_options(*args):
            option = self.search_option_var.get()
            if option == "Profesional":
                self.search_entry.update_options(profesionales)
            elif option == "Comitente":
                self.search_entry.update_options(comitentes)
            else:
                self.search_entry.update_options([])  # Sin autocompletado para otros campos
        
        # Vincular el cambio de opción de búsqueda
        self.search_option_var.trace_add("write", update_search_options)
        
        # Botón de búsqueda
        search_button = ctk.CTkButton(search_frame, text="Buscar", command=self.search_obras)
        search_button.grid(row=0, column=3, padx=5, pady=5)
        
        # Configurar grid para ser responsive
        search_frame.grid_columnconfigure(2, weight=1)
        
        # Obtener todas las obras
        self.obras = self.data_manager.get_all_works("obra")
        
        # Crear lista de obras para mostrar
        obras_display = []
        for obra in self.obras:
            display = f"{obra['nombre_profesional']} - {obra['nombre_comitente']} ({obra['fecha']})"
            obras_display.append(display)
        
        # Selector de obra
        selection_menu_frame = ctk.CTkFrame(selection_frame)
        selection_menu_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ctk.CTkLabel(selection_menu_frame, text="Seleccionar Obra:").pack(side=tk.LEFT, padx=5)
        
        # Usamos una variable separada para el texto mostrado
        self.obra_selector_var = ctk.StringVar(value="Seleccione una obra" if obras_display else "No hay obras registradas")
        
        self.obra_selector = ctk.CTkOptionMenu(
            selection_menu_frame, 
            values=obras_display if obras_display else ["No hay obras registradas"],
            variable=self.obra_selector_var,
            command=self.on_obra_selected,
            dynamic_resizing=False,
            width=300
        )
        self.obra_selector.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Frame para los datos de la obra seleccionada
        self.obra_edit_frame = ctk.CTkScrollableFrame(parent)
        self.obra_edit_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Variables para los campos editables
        self.edit_obra_vars = {
            "tasa_sellado": ctk.StringVar(),
            "tasa_visado": ctk.StringVar(),
            "estado_pago_sellado": ctk.StringVar(),
            "estado_pago_visado": ctk.StringVar(),
            "nro_expediente_cpim": ctk.StringVar(),
            "fecha_salida": ctk.StringVar(),
            "persona_retira": ctk.StringVar(),
            "nro_caja": ctk.StringVar()
        }
        
        # Indicar que no hay obra seleccionada
        ctk.CTkLabel(self.obra_edit_frame, text="Seleccione una obra para editar", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
        
    def search_obras(self):
        """Busca obras según el criterio seleccionado"""
        search_text = self.search_entry.get().strip().lower()
        search_option = self.search_option_var.get()
        
        if not search_text:
            messagebox.showwarning("Búsqueda vacía", "Por favor ingrese un texto para buscar")
            return
        
        # Obtener todas las obras de nuevo (por si se agregaron nuevas)
        self.obras = self.data_manager.get_all_works("obra")
        
        # Filtrar según el criterio
        filtered_obras = []
        
        for obra in self.obras:
            # Obtener obra completa para acceder a todos los campos
            obra_completa = self.data_manager.get_work_by_id("obra", obra["id"])
            
            if obra_completa:
                if search_option == "Profesional" and obra_completa["nombre_profesional"] and search_text in obra_completa["nombre_profesional"].lower():
                    filtered_obras.append(obra)
                elif search_option == "Comitente" and obra_completa["nombre_comitente"] and search_text in obra_completa["nombre_comitente"].lower():
                    filtered_obras.append(obra)
                elif search_option == "Partida Inmobiliaria" and obra_completa["nro_partida_inmobiliaria"] and search_text in str(obra_completa["nro_partida_inmobiliaria"]).lower():
                    filtered_obras.append(obra)
                elif search_option == "Número GOP" and obra_completa["nro_sistema_gop"] and search_text in str(obra_completa["nro_sistema_gop"]).lower():
                    filtered_obras.append(obra)
        
        # Actualizar la lista de obras mostradas
        if filtered_obras:
            obras_display = []
            for obra in filtered_obras:
                display = f"{obra['nombre_profesional']} - {obra['nombre_comitente']} ({obra['fecha']})"
                obras_display.append(display)
            
            # Guardar la lista filtrada
            self.obras = filtered_obras
            
            # Actualizar el selector
            self.obra_selector.configure(values=obras_display)
            self.obra_selector_var.set(obras_display[0])
            
            # Seleccionar la primera obra
            self.on_obra_selected(obras_display[0])
        else:
            messagebox.showinfo("Búsqueda", f"No se encontraron obras que coincidan con '{search_text}' en {search_option}")

    
    def on_obra_selected(self, selection):
        """Maneja la selección de una obra para editar"""
        # Limpiar el frame de edición
        for widget in self.obra_edit_frame.winfo_children():
            widget.destroy()
        
        # Obtener el índice de la obra seleccionada
        try:
            if not selection or selection == "No hay obras registradas":
                ctk.CTkLabel(self.obra_edit_frame, text="No hay obras disponibles para editar", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
                return
                
            index = self.obra_selector.cget("values").index(selection)
            obra_id = self.obras[index]["id"]
            
            # Obtener los datos completos de la obra
            obra = self.data_manager.get_work_by_id("obra", obra_id)
            
            if obra:
                # Mostrar información no editable
                info_frame = ctk.CTkFrame(self.obra_edit_frame)
                info_frame.pack(fill=tk.X, padx=10, pady=10)
                
                info_text = f"Profesional: {obra['nombre_profesional']}\n"
                info_text += f"Comitente: {obra['nombre_comitente']}\n"
                info_text += f"Tipo: {obra['tipo_trabajo']}\n"
                info_text += f"Ubicación: {obra['ubicacion']}"
                
                ctk.CTkLabel(info_frame, text=info_text, justify=tk.LEFT).pack(pady=10, padx=10, side=tk.LEFT)
                
                # Botón para repetir trabajo con otro profesional
                btn_repeat = ctk.CTkButton(info_frame, text="Repetir con otro Profesional", 
                                        command=lambda: self.repeat_obra_with_new_professional(obra_id))
                btn_repeat.pack(pady=10, padx=10, side=tk.RIGHT)
                
                # Crear campos editables
                edit_frame = ctk.CTkFrame(self.obra_edit_frame)
                edit_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                
                row = 0
                
                # Crear widgets temporales separados de las variables
                # Tasa de sellado
                ctk.CTkLabel(edit_frame, text="Tasa de Sellado:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                sellado_entry = ctk.CTkEntry(edit_frame)
                sellado_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                sellado_entry.insert(0, obra["tasa_sellado"] if obra["tasa_sellado"] else "")
                self.edit_obra_vars["tasa_sellado"] = sellado_entry
                row += 1
                
                # Añadir un separador o título para la sección de visados específicos
                ctk.CTkLabel(edit_frame, text="Visados Específicos", font=("Arial", 12, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", pady=(15, 5), padx=5)
                row += 1
                
                # Visado de instalación de Gas
                ctk.CTkLabel(edit_frame, text="Visado de instalación de Gas:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                visado_gas_entry = ctk.CTkEntry(edit_frame)
                visado_gas_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                visado_gas_entry.insert(0, obra["visado_gas"] if obra["visado_gas"] else "")
                self.edit_obra_vars["visado_gas"] = visado_gas_entry
                row += 1
                
                # Visado de instalación de Salubridad
                ctk.CTkLabel(edit_frame, text="Visado de instalación de Salubridad:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                visado_salubridad_entry = ctk.CTkEntry(edit_frame)
                visado_salubridad_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                visado_salubridad_entry.insert(0, obra["visado_salubridad"] if obra["visado_salubridad"] else "")
                self.edit_obra_vars["visado_salubridad"] = visado_salubridad_entry
                row += 1
                
                # Visado de instalación eléctrica
                ctk.CTkLabel(edit_frame, text="Visado de instalación eléctrica:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                visado_electrica_entry = ctk.CTkEntry(edit_frame)
                visado_electrica_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                visado_electrica_entry.insert(0, obra["visado_electrica"] if obra["visado_electrica"] else "")
                self.edit_obra_vars["visado_electrica"] = visado_electrica_entry
                row += 1
                
                # Visado de instalación electromecánica
                ctk.CTkLabel(edit_frame, text="Visado de instalación electromecánica:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                visado_electromecanica_entry = ctk.CTkEntry(edit_frame)
                visado_electromecanica_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                visado_electromecanica_entry.insert(0, obra["visado_electromecanica"] if obra["visado_electromecanica"] else "")
                self.edit_obra_vars["visado_electromecanica"] = visado_electromecanica_entry
                row += 1
                
                # Estado de pago sellado
                ctk.CTkLabel(edit_frame, text="Estado Pago Sellado:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                estado_frame = ctk.CTkFrame(edit_frame)
                estado_frame.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                
                self.estado_sellado_var = ctk.StringVar(value=obra["estado_pago_sellado"] if obra["estado_pago_sellado"] else "No pagado")
                ctk.CTkRadioButton(estado_frame, text="Pagado", variable=self.estado_sellado_var, 
                                value="Pagado").pack(side=tk.LEFT, padx=10)
                ctk.CTkRadioButton(estado_frame, text="No pagado", variable=self.estado_sellado_var, 
                                value="No pagado").pack(side=tk.LEFT, padx=10)
                self.edit_obra_vars["estado_pago_sellado"] = self.estado_sellado_var
                row += 1
                
                # Estado de pago visado
                ctk.CTkLabel(edit_frame, text="Estado Pago Visado:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                estado_frame = ctk.CTkFrame(edit_frame)
                estado_frame.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                
                self.estado_visado_var = ctk.StringVar(value=obra["estado_pago_visado"] if obra["estado_pago_visado"] else "No pagado")
                ctk.CTkRadioButton(estado_frame, text="Pagado", variable=self.estado_visado_var, 
                                value="Pagado").pack(side=tk.LEFT, padx=10)
                ctk.CTkRadioButton(estado_frame, text="No pagado", variable=self.estado_visado_var, 
                                value="No pagado").pack(side=tk.LEFT, padx=10)
                self.edit_obra_vars["estado_pago_visado"] = self.estado_visado_var
                row += 1
                
                # Nro. de expediente CPIM
                ctk.CTkLabel(edit_frame, text="Nro. Expediente CPIM:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                cpim_entry = ctk.CTkEntry(edit_frame)
                cpim_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                cpim_entry.insert(0, obra["nro_expediente_cpim"] if obra["nro_expediente_cpim"] else "")
                self.edit_obra_vars["nro_expediente_cpim"] = cpim_entry
                row += 1
                
                # Fecha de salida
                ctk.CTkLabel(edit_frame, text="Fecha de Salida:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                fecha_entry = ctk.CTkEntry(edit_frame)
                fecha_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                fecha_entry.insert(0, obra["fecha_salida"] if obra["fecha_salida"] else "")
                self.edit_obra_vars["fecha_salida"] = fecha_entry
                row += 1
                
                # Persona que retira
                ctk.CTkLabel(edit_frame, text="Persona que Retira:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                persona_entry = ctk.CTkEntry(edit_frame)
                persona_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                persona_entry.insert(0, obra["persona_retira"] if obra["persona_retira"] else "")
                self.edit_obra_vars["persona_retira"] = persona_entry
                row += 1
                
                # Nro. de caja
                ctk.CTkLabel(edit_frame, text="Nro. de Caja:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                caja_entry = ctk.CTkEntry(edit_frame)
                caja_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                caja_entry.insert(0, obra["nro_caja"] if obra["nro_caja"] else "")
                self.edit_obra_vars["nro_caja"] = caja_entry
                row += 1
                
                # Si es formato digital, mostrar botón para abrir carpeta
                if obra["formato"] == "Digital" and obra["ruta_carpeta"]:
                    btn_open_folder = ctk.CTkButton(edit_frame, text="Abrir Carpeta", 
                                                command=lambda: self.file_manager.open_folder(obra["ruta_carpeta"]))
                    btn_open_folder.grid(row=row, column=0, pady=10, padx=5)
                    row += 1
                
                # Configurar el grid para que sea responsive
                edit_frame.grid_columnconfigure(1, weight=1)
                
                # Botón para guardar cambios
                btn_save = ctk.CTkButton(edit_frame, text="Guardar Cambios", 
                                        font=ctk.CTkFont(size=14),
                                        command=lambda: self.save_obra_changes(obra_id))
                btn_save.grid(row=row, column=0, columnspan=2, pady=20, padx=5, sticky="ew")
            else:
                ctk.CTkLabel(self.obra_edit_frame, text="No se pudo cargar la información de la obra", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
        
        except Exception as e:
            print(f"Error al cargar la obra: {e}")
            ctk.CTkLabel(self.obra_edit_frame, text=f"Error al cargar la obra: {str(e)}", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
    
    def save_obra_changes(self, obra_id):
        """Guarda los cambios realizados a una obra"""
        try:
            # Obtener datos del formulario
            data = {}
            
            # Para entradas normales
            for key, widget in self.edit_obra_vars.items():
                if key in ["estado_pago_sellado", "estado_pago_visado"]:
                    # Para StringVars (radiobuttons)
                    data[key] = widget.get()
                else:
                    # Para widgets Entry
                    data[key] = widget.get()
            
            # Guardar cambios
            if self.data_manager.update_obra_general(obra_id, data):
                # Determinar si son datos de salida
                campos_salida = ["estado_pago_sellado", "estado_pago_visado", "nro_expediente_cpim", "fecha_salida", "persona_retira", "nro_caja"]
                es_actualizacion_salida = any(campo in data for campo in campos_salida)
                
                if es_actualizacion_salida:
                    messagebox.showinfo("Éxito", "Cambios guardados correctamente. Si existen trabajos similares de otros profesionales, también se han actualizado sus datos de salida.")
                else:
                    messagebox.showinfo("Éxito", "Cambios guardados correctamente.")
            else:
                messagebox.showerror("Error", "No se pudieron guardar los cambios.")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar cambios: {str(e)}")

    def repeat_obra_with_new_professional(self, obra_id):
        """Prepara un nuevo registro basado en una obra existente pero para otro profesional"""
        # Obtener los datos completos de la obra
        obra = self.data_manager.get_work_by_id("obra", obra_id)
        
        if not obra:
            messagebox.showerror("Error", "No se pudo cargar la información de la obra original")
            return
        
        # Cambiar a la ventana de nuevo registro
        self.show_new_record_window()
        
        # Seleccionar la pestaña de "Obra en general"
        try:
            # Acceder a las pestañas (esto puede variar según cómo esté implementada tu interfaz)
            for widget in self.winfo_children():
                if isinstance(widget, ctk.CTkFrame):
                    for child in widget.winfo_children():
                        if isinstance(child, ctk.CTkTabview):
                            # Seleccionar la pestaña de Obra en general
                            child.set("Obra en general")
                            break
        except Exception as e:
            print(f"Error al cambiar a la pestaña de Obra en general: {e}")
        
        # Copiar los datos relevantes al formulario de nueva obra, excepto los relacionados con el profesional
        try:
            # Formato
            self.formato_var.set(obra["formato"])
            self.toggle_formato_fields()
            
            # Números de copias o GOP según formato
            if obra["formato"] == "Físico":
                self.obra_vars["nro_copias"].set(obra["nro_copias"])
            else:
                self.obra_vars["nro_sistema_gop"].set(obra["nro_sistema_gop"])
            
            # Tipo de trabajo
            self.obra_vars["tipo_trabajo"].set(obra["tipo_trabajo"])
            
            # Datos del comitente y ubicación
            self.obra_vars["nombre_comitente"].set(obra["nombre_comitente"])
            # Forzar que se oculte el dropdown de autocompletado
            self.comit_entry.hide_dropdown()
            
            self.obra_vars["ubicacion"].set(obra["ubicacion"])
            self.obra_vars["nro_expte_municipal"].set(obra["nro_expte_municipal"])
            self.obra_vars["nro_partida_inmobiliaria"].set(obra["nro_partida_inmobiliaria"])
            
            # Tasas
            self.obra_vars["tasa_sellado"].set(obra["tasa_sellado"])
            self.obra_vars["tasa_visado"].set(obra["tasa_visado"])
            
            # Nuevos campos de visado
            self.obra_vars["visado_gas"].set(obra["visado_gas"])
            self.obra_vars["visado_salubridad"].set(obra["visado_salubridad"])
            self.obra_vars["visado_electrica"].set(obra["visado_electrica"])
            self.obra_vars["visado_electromecanica"].set(obra["visado_electromecanica"])
            
            # Limpiar específicamente los datos del profesional
            self.obra_vars["profesion"].set("")
            self.obra_vars["nombre_profesional"].set("")
            
            # Mostrar mensaje informativo
            messagebox.showinfo("Repetir trabajo", 
                            "Se han copiado los datos del trabajo seleccionado.\n\n"
                            "Por favor, complete la información del nuevo profesional y guarde el registro.")
        
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron copiar todos los datos: {str(e)}")

    def create_main_menu(self):
        """Crea el menú principal de la aplicación"""
        # Limpiar la ventana
        for widget in self.winfo_children():
            widget.destroy()
        
        # Crear frame principal
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        title = ctk.CTkLabel(main_frame, text="Sistema de Gestión - Consejo Profesional de Ingeniería de Misiones", 
                            font=ctk.CTkFont(size=24, weight="bold"))
        title.pack(pady=20)
        
        # Botones
        btn_frame = ctk.CTkFrame(main_frame)
        btn_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Botón para nuevo registro
        btn_new = ctk.CTkButton(btn_frame, text="Nuevo Registro", 
                            font=ctk.CTkFont(size=16),
                            height=80,
                            command=self.show_new_record_window)
        btn_new.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        # Botón para editar trabajo
        btn_edit = ctk.CTkButton(btn_frame, text="Editar Trabajo", 
                                font=ctk.CTkFont(size=16),
                                height=80,
                                command=self.show_edit_work_window)
        btn_edit.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        
        # Botón para duplicar trabajo
        btn_duplicate = ctk.CTkButton(btn_frame, text="Repetir Trabajo\ncon Otro Profesional", 
                                    font=ctk.CTkFont(size=16),
                                    height=80,
                                    command=self.show_duplicate_work_window)
        btn_duplicate.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")
        
        # Botón para generar documento Word
        btn_word = ctk.CTkButton(btn_frame, text="Generar Documento Word", 
                                font=ctk.CTkFont(size=16),
                                height=80,
                                command=self.show_generate_word_window)
        btn_word.grid(row=1, column=1, padx=20, pady=20, sticky="nsew")
        
        # Botón para ver carpeta de archivos
        btn_folder = ctk.CTkButton(btn_frame, text="Ver Carpeta de Archivos", 
                                font=ctk.CTkFont(size=16),
                                height=80,
                                command=lambda: self.file_manager.open_folder(TRABAJOS_PATH))
        btn_folder.grid(row=2, column=0, columnspan=2, padx=20, pady=20, sticky="nsew")
        
        # Configurar el grid para que sea responsive
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)
        btn_frame.grid_rowconfigure(0, weight=1)
        btn_frame.grid_rowconfigure(1, weight=1)
        btn_frame.grid_rowconfigure(2, weight=1)

    def show_duplicate_work_window(self):
        """Muestra la ventana para seleccionar un trabajo a duplicar"""
        # Limpiar la ventana
        for widget in self.winfo_children():
            widget.destroy()
        
        # Crear frame principal
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        title = ctk.CTkLabel(main_frame, text="Repetir Trabajo con Otro Profesional", 
                            font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=10)
        
        # Frame para búsqueda
        search_frame = ctk.CTkFrame(main_frame)
        search_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Etiqueta para búsqueda
        ctk.CTkLabel(search_frame, text="Buscar por:").grid(row=0, column=0, padx=5, pady=5)
        
        # Opciones de búsqueda
        search_options = ["Profesional", "Comitente", "Partida Inmobiliaria", "Número GOP"]
        self.duplicate_search_option_var = ctk.StringVar(value=search_options[0])
        option_menu = ctk.CTkOptionMenu(search_frame, values=search_options, variable=self.duplicate_search_option_var)
        option_menu.grid(row=0, column=1, padx=5, pady=5)
        
        # Obtener listas de profesionales y comitentes existentes
        profesionales = self.data_manager.get_all_profesionales()
        comitentes = self.data_manager.get_all_comitentes()
        
        # Campo de búsqueda con autocompletado
        self.duplicate_search_entry = AutocompleteEntry(search_frame, width=200, options=profesionales)
        self.duplicate_search_entry.grid(row=0, column=2, padx=5, pady=5)
        
        # Actualizar opciones de autocompletado según el criterio seleccionado
        def update_search_options(*args):
            option = self.duplicate_search_option_var.get()
            if option == "Profesional":
                self.duplicate_search_entry.update_options(profesionales)
            elif option == "Comitente":
                self.duplicate_search_entry.update_options(comitentes)
            else:
                self.duplicate_search_entry.update_options([])  # Sin autocompletado para otros campos
        
        # Vincular el cambio de opción de búsqueda
        self.duplicate_search_option_var.trace_add("write", update_search_options)
        
        # Botón de búsqueda
        search_button = ctk.CTkButton(search_frame, text="Buscar", command=self.search_obras_for_duplication)
        search_button.grid(row=0, column=3, padx=5, pady=5)
        
        # Configurar grid para ser responsive
        search_frame.grid_columnconfigure(2, weight=1)
        
        # Obtener todas las obras
        self.duplicate_obras = self.data_manager.get_all_works("obra")
        
        # Crear lista de obras para mostrar
        obras_display = []
        for obra in self.duplicate_obras:
            display = f"{obra['nombre_profesional']} - {obra['nombre_comitente']} ({obra['fecha']})"
            obras_display.append(display)
        
        # Selector de obra
        selection_frame = ctk.CTkFrame(main_frame)
        selection_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ctk.CTkLabel(selection_frame, text="Seleccionar Obra:").pack(side=tk.LEFT, padx=5)
        
        # Usamos una variable separada para el texto mostrado
        self.duplicate_obra_selector_var = ctk.StringVar(value="Seleccione una obra" if obras_display else "No hay obras registradas")
        
        self.duplicate_obra_selector = ctk.CTkOptionMenu(
            selection_frame, 
            values=obras_display if obras_display else ["No hay obras registradas"],
            variable=self.duplicate_obra_selector_var,
            command=self.on_duplicate_obra_selected,
            dynamic_resizing=False,
            width=300
        )
        self.duplicate_obra_selector.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Frame para mostrar la información de la obra seleccionada
        self.duplicate_obra_frame = ctk.CTkScrollableFrame(main_frame)
        self.duplicate_obra_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Indicar que se seleccione una obra
        ctk.CTkLabel(self.duplicate_obra_frame, text="Seleccione una obra para duplicar", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
        
        # Botón para volver al menú principal
        btn_back = ctk.CTkButton(main_frame, text="Volver al Menú Principal", 
                                font=ctk.CTkFont(size=14),
                                command=self.create_main_menu)
        btn_back.pack(pady=10)

    def search_obras_for_duplication(self):
        """Busca obras para duplicar según el criterio seleccionado"""
        search_text = self.duplicate_search_entry.get().strip().lower()
        search_option = self.duplicate_search_option_var.get()
        
        if not search_text:
            messagebox.showwarning("Búsqueda vacía", "Por favor ingrese un texto para buscar")
            return
        
        # Obtener todas las obras de nuevo (por si se agregaron nuevas)
        self.duplicate_obras = self.data_manager.get_all_works("obra")
        
        # Filtrar según el criterio
        filtered_obras = []
        
        for obra in self.duplicate_obras:
            # Obtener obra completa para acceder a todos los campos
            obra_completa = self.data_manager.get_work_by_id("obra", obra["id"])
            
            if obra_completa:
                if search_option == "Profesional" and obra_completa["nombre_profesional"] and search_text in obra_completa["nombre_profesional"].lower():
                    filtered_obras.append(obra)
                elif search_option == "Comitente" and obra_completa["nombre_comitente"] and search_text in obra_completa["nombre_comitente"].lower():
                    filtered_obras.append(obra)
                elif search_option == "Partida Inmobiliaria" and obra_completa["nro_partida_inmobiliaria"] and search_text in str(obra_completa["nro_partida_inmobiliaria"]).lower():
                    filtered_obras.append(obra)
                elif search_option == "Número GOP" and obra_completa["nro_sistema_gop"] and search_text in str(obra_completa["nro_sistema_gop"]).lower():
                    filtered_obras.append(obra)
        
        # Actualizar la lista de obras mostradas
        if filtered_obras:
            obras_display = []
            for obra in filtered_obras:
                display = f"{obra['nombre_profesional']} - {obra['nombre_comitente']} ({obra['fecha']})"
                obras_display.append(display)
            
            # Guardar la lista filtrada
            self.duplicate_obras = filtered_obras
            
            # Actualizar el selector
            self.duplicate_obra_selector.configure(values=obras_display)
            self.duplicate_obra_selector_var.set(obras_display[0])
            
            # Seleccionar la primera obra
            self.on_duplicate_obra_selected(obras_display[0])
        else:
            messagebox.showinfo("Búsqueda", f"No se encontraron obras que coincidan con '{search_text}' en {search_option}")

    def on_duplicate_obra_selected(self, selection):
        """Maneja la selección de una obra para duplicar"""
        # Limpiar el frame de información
        for widget in self.duplicate_obra_frame.winfo_children():
            widget.destroy()
        
        # Obtener el índice de la obra seleccionada
        try:
            if not selection or selection == "No hay obras registradas":
                ctk.CTkLabel(self.duplicate_obra_frame, text="No hay obras disponibles para duplicar", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
                return
                
            index = self.duplicate_obra_selector.cget("values").index(selection)
            obra_id = self.duplicate_obras[index]["id"]
            
            # Obtener los datos completos de la obra
            obra = self.data_manager.get_work_by_id("obra", obra_id)
            
            if obra:
                # Mostrar información detallada de la obra
                info_frame = ctk.CTkFrame(self.duplicate_obra_frame)
                info_frame.pack(fill=tk.X, padx=10, pady=10)
                
                info_text = f"Profesional: {obra['nombre_profesional']}\n"
                info_text += f"Profesión: {obra['profesion']}\n"
                info_text += f"Comitente: {obra['nombre_comitente']}\n"
                info_text += f"Ubicación: {obra['ubicacion']}\n"
                info_text += f"Tipo: {obra['tipo_trabajo']}\n"
                info_text += f"Formato: {obra['formato']}\n"
                if obra['formato'] == "Físico":
                    info_text += f"Nro. Copias: {obra['nro_copias']}\n"
                else:
                    info_text += f"Nro. Sistema GOP: {obra['nro_sistema_gop']}\n"
                info_text += f"Nro. Expte. Municipal: {obra['nro_expte_municipal']}\n"
                info_text += f"Nro. Partida Inmobiliaria: {obra['nro_partida_inmobiliaria']}\n"
                info_text += f"Tasa Sellado: {obra['tasa_sellado']}\n"
                info_text += f"Tasa Visado: {obra['tasa_visado']}"
                
                ctk.CTkLabel(info_frame, text=info_text, justify=tk.LEFT).pack(pady=10, padx=10)
                
                # Botón para duplicar esta obra
                btn_frame = ctk.CTkFrame(self.duplicate_obra_frame)
                btn_frame.pack(fill=tk.X, padx=10, pady=10)
                
                btn_duplicate = ctk.CTkButton(btn_frame, text="Duplicar con Nuevo Profesional", 
                                            font=ctk.CTkFont(size=14),
                                            command=lambda: self.repeat_obra_with_new_professional(obra_id))
                btn_duplicate.pack(pady=10, padx=10, fill=tk.X)
                
                # Instrucciones
                instructions = "Al hacer clic en 'Duplicar con Nuevo Profesional', se abrirá el formulario de nuevo registro con todos los datos de este trabajo pre-cargados, excepto los datos del profesional.\n\nSolo tendrá que completar la información del nuevo profesional para registrar el trabajo."
                ctk.CTkLabel(self.duplicate_obra_frame, text=instructions, wraplength=600, justify=tk.LEFT).pack(pady=20, padx=10)
            
            else:
                ctk.CTkLabel(self.duplicate_obra_frame, text="No se pudo cargar la información de la obra", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
        
        except Exception as e:
            print(f"Error al cargar la obra para duplicar: {e}")
            ctk.CTkLabel(self.duplicate_obra_frame, text=f"Error al cargar la obra: {str(e)}", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
    
    def setup_edit_informe_tab(self, parent):
        """Configura la pestaña para editar informes técnicos"""
        # Crear frame superior para búsqueda y selección
        selection_frame = ctk.CTkFrame(parent)
        selection_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Frame para búsqueda
        search_frame = ctk.CTkFrame(selection_frame)
        search_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Etiqueta para búsqueda
        ctk.CTkLabel(search_frame, text="Buscar por:").grid(row=0, column=0, padx=5, pady=5)
        
        # Opciones de búsqueda
        search_options = ["Profesional", "Comitente", "Tipo de Informe"]
        self.informe_search_option_var = ctk.StringVar(value=search_options[0])
        option_menu = ctk.CTkOptionMenu(search_frame, values=search_options, variable=self.informe_search_option_var)
        option_menu.grid(row=0, column=1, padx=5, pady=5)
        
        # Obtener listas de profesionales y comitentes existentes
        profesionales = self.data_manager.get_all_profesionales()
        comitentes = self.data_manager.get_all_comitentes()
        
        # Campo de búsqueda con autocompletado
        self.informe_search_entry = AutocompleteEntry(search_frame, width=200, options=profesionales)
        self.informe_search_entry.grid(row=0, column=2, padx=5, pady=5)
        
        # Actualizar opciones de autocompletado según el criterio seleccionado
        def update_search_options(*args):
            option = self.informe_search_option_var.get()
            if option == "Profesional":
                self.informe_search_entry.update_options(profesionales)
            elif option == "Comitente":
                self.informe_search_entry.update_options(comitentes)
            else:
                self.informe_search_entry.update_options([])  # Sin autocompletado para otros campos
        
        # Vincular el cambio de opción de búsqueda
        self.informe_search_option_var.trace_add("write", update_search_options)
        
        # Botón de búsqueda
        search_button = ctk.CTkButton(search_frame, text="Buscar", command=self.search_informes)
        search_button.grid(row=0, column=3, padx=5, pady=5)
        
        # Configurar grid para ser responsive
        search_frame.grid_columnconfigure(2, weight=1)
        
        # Obtener todos los informes
        self.informes = self.data_manager.get_all_works("informe")
        
        # Crear lista de informes para mostrar
        informes_display = []
        for informe in self.informes:
            display = f"{informe['profesional']} - {informe['comitente']} ({informe['fecha']})"
            informes_display.append(display)
        
        # Selector de informe
        selection_menu_frame = ctk.CTkFrame(selection_frame)
        selection_menu_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ctk.CTkLabel(selection_menu_frame, text="Seleccionar Informe:").pack(side=tk.LEFT, padx=5)
        
        # Usamos una variable separada para el texto mostrado
        self.informe_selector_var = ctk.StringVar(value="Seleccione un informe" if informes_display else "No hay informes registrados")
        
        self.informe_selector = ctk.CTkOptionMenu(
            selection_menu_frame, 
            values=informes_display if informes_display else ["No hay informes registrados"],
            variable=self.informe_selector_var,
            command=self.on_informe_selected,
            dynamic_resizing=False,
            width=300
        )
        self.informe_selector.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Frame para los datos del informe seleccionado
        self.informe_edit_frame = ctk.CTkScrollableFrame(parent)
        self.informe_edit_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Variables para los campos editables (ahora usaremos widgets en lugar de StringVars)
        self.edit_informe_vars = {}
        
        # Indicar que no hay informe seleccionado
        ctk.CTkLabel(self.informe_edit_frame, text="Seleccione un informe para editar", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
        
        # Seleccionar el primer informe si hay alguno
        if informes_display:
            self.informe_selector.set(informes_display[0])
            self.on_informe_selected(informes_display[0])
        
    def search_informes(self):
        """Busca informes según el criterio seleccionado"""
        # Verifica que la variable de búsqueda exista
        if not hasattr(self, 'informe_search_entry'):
            print("No se encontró informe_search_entry, cambiando a usar otra variable...")
            # Puedes buscar otra variable que podría estar disponible
            # Por ejemplo, si estás usando informe_search_var o alguna otra
            search_text = self.informe_search_var.get().strip().lower() if hasattr(self, 'informe_search_var') else ""
        else:
            search_text = self.informe_search_entry.get().strip().lower()
        
        # También verificamos la variable de opción de búsqueda
        if not hasattr(self, 'informe_search_option_var'):
            print("No se encontró informe_search_option_var...")
            search_option = "Profesional"  # valor predeterminado
        else:
            search_option = self.informe_search_option_var.get()
        
        if not search_text:
            messagebox.showwarning("Búsqueda vacía", "Por favor ingrese un texto para buscar")
            return
        
        # Obtener todos los informes de nuevo (por si se agregaron nuevos)
        self.informes = self.data_manager.get_all_works("informe")
        
        # Filtrar según el criterio
        filtered_informes = []
        
        for informe in self.informes:
            # Obtener informe completo para acceder a todos los campos
            informe_completo = self.data_manager.get_work_by_id("informe", informe["id"])
            
            if informe_completo:
                if search_option == "Profesional" and informe_completo["profesional"] and search_text in informe_completo["profesional"].lower():
                    filtered_informes.append(informe)
                elif search_option == "Comitente" and informe_completo["comitente"] and search_text in informe_completo["comitente"].lower():
                    filtered_informes.append(informe)
                elif search_option == "Tipo de Informe" and informe_completo["tipo_trabajo"] and search_text in str(informe_completo["tipo_trabajo"]).lower():
                    filtered_informes.append(informe)
        
        # Actualizar la lista de informes mostrados
        if filtered_informes:
            informes_display = []
            for informe in filtered_informes:
                display = f"{informe['profesional']} - {informe['comitente']} ({informe['fecha']})"
                informes_display.append(display)
            
            # Guardar la lista filtrada
            self.informes = filtered_informes
            
            # Actualizar el selector
            self.informe_selector.configure(values=informes_display)
            self.informe_selector_var.set(informes_display[0])
            
            # Seleccionar el primer informe
            self.on_informe_selected(informes_display[0])
        else:
            messagebox.showinfo("Búsqueda", f"No se encontraron informes que coincidan con '{search_text}' en {search_option}")
    
    def on_informe_selected(self, selection):
        """Maneja la selección de un informe para editar"""
        # Limpiar el frame de edición
        for widget in self.informe_edit_frame.winfo_children():
            widget.destroy()
        
        # Obtener el índice del informe seleccionado
        try:
            if not selection or selection == "No hay informes registrados":
                ctk.CTkLabel(self.informe_edit_frame, text="No hay informes disponibles para editar", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
                return
                
            index = self.informe_selector.cget("values").index(selection)
            informe_id = self.informes[index]["id"]
            
            # Obtener los datos completos del informe
            informe = self.data_manager.get_work_by_id("informe", informe_id)
            
            if informe:
                # Mostrar información no editable
                info_frame = ctk.CTkFrame(self.informe_edit_frame)
                info_frame.pack(fill=tk.X, padx=10, pady=10)
                
                info_text = f"Profesional: {informe['profesional']}\n"
                info_text += f"Comitente: {informe['comitente']}\n"
                info_text += f"Tipo: {informe['tipo_trabajo']}\n"
                info_text += f"Detalle: {informe['detalle']}"
                
                ctk.CTkLabel(info_frame, text=info_text, justify=tk.LEFT).pack(pady=10, padx=10)
                
                # Crear campos editables
                edit_frame = ctk.CTkFrame(self.informe_edit_frame)
                edit_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                
                row = 0
                
                # Tasa de sellado
                ctk.CTkLabel(edit_frame, text="Tasa de Sellado:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                sellado_entry = ctk.CTkEntry(edit_frame)
                sellado_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                sellado_entry.insert(0, informe["tasa_sellado"] if informe["tasa_sellado"] else "")
                self.edit_informe_vars["tasa_sellado"] = sellado_entry
                row += 1
                
                # Estado de pago
                ctk.CTkLabel(edit_frame, text="Estado de Pago:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                estado_frame = ctk.CTkFrame(edit_frame)
                estado_frame.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                
                self.estado_pago_var = ctk.StringVar(value=informe["estado_pago"] if informe["estado_pago"] else "No pagado")
                ctk.CTkRadioButton(estado_frame, text="Pagado", variable=self.estado_pago_var, 
                                value="Pagado").pack(side=tk.LEFT, padx=10)
                ctk.CTkRadioButton(estado_frame, text="No pagado", variable=self.estado_pago_var, 
                                value="No pagado").pack(side=tk.LEFT, padx=10)
                self.edit_informe_vars["estado_pago"] = self.estado_pago_var
                row += 1
                
                # Nro. de expediente CPIM
                ctk.CTkLabel(edit_frame, text="Nro. Expediente CPIM:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                cpim_entry = ctk.CTkEntry(edit_frame)
                cpim_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                cpim_entry.insert(0, informe["nro_expediente_cpim"] if informe["nro_expediente_cpim"] else "")
                self.edit_informe_vars["nro_expediente_cpim"] = cpim_entry
                row += 1
                
                # Fecha de salida
                ctk.CTkLabel(edit_frame, text="Fecha de Salida:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                fecha_entry = ctk.CTkEntry(edit_frame)
                fecha_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                fecha_entry.insert(0, informe["fecha_salida"] if informe["fecha_salida"] else "")
                self.edit_informe_vars["fecha_salida"] = fecha_entry
                row += 1
                
                # Persona que retira
                ctk.CTkLabel(edit_frame, text="Persona que Retira:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                persona_entry = ctk.CTkEntry(edit_frame)
                persona_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                persona_entry.insert(0, informe["persona_retira"] if informe["persona_retira"] else "")
                self.edit_informe_vars["persona_retira"] = persona_entry
                row += 1
                
                # Nro. de caja
                ctk.CTkLabel(edit_frame, text="Nro. de Caja:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
                caja_entry = ctk.CTkEntry(edit_frame)
                caja_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                caja_entry.insert(0, informe["nro_caja"] if informe["nro_caja"] else "")
                self.edit_informe_vars["nro_caja"] = caja_entry
                row += 1
                
                # Si es formato digital, mostrar botón para abrir carpeta
                if informe["formato"] == "Digital" and informe["ruta_carpeta"]:
                    btn_open_folder = ctk.CTkButton(edit_frame, text="Abrir Carpeta", 
                                                command=lambda: self.file_manager.open_folder(informe["ruta_carpeta"]))
                    btn_open_folder.grid(row=row, column=0, pady=10, padx=5)
                    row += 1
                    
                # Configurar el grid para que sea responsive
                edit_frame.grid_columnconfigure(1, weight=1)
                
                # Botón para guardar cambios
                btn_save = ctk.CTkButton(edit_frame, text="Guardar Cambios", 
                                        font=ctk.CTkFont(size=14),
                                        command=lambda: self.save_informe_changes(informe_id))
                btn_save.grid(row=row, column=0, columnspan=2, pady=20, padx=5, sticky="ew")
            else:
                ctk.CTkLabel(self.informe_edit_frame, text="No se pudo cargar la información del informe", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
        
        except Exception as e:
            print(f"Error al cargar el informe: {e}")
            ctk.CTkLabel(self.informe_edit_frame, text=f"Error al cargar el informe: {str(e)}", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
            
    def save_informe_changes(self, informe_id):
        """Guarda los cambios realizados a un informe"""
        try:
            # Obtener datos del formulario
            data = {}
            
            # Para entradas normales
            for key, widget in self.edit_informe_vars.items():
                if key in ["estado_pago"]:
                    # Para StringVars (radiobuttons)
                    data[key] = widget.get()
                else:
                    # Para widgets Entry
                    data[key] = widget.get()
            
            # Guardar cambios
            if self.data_manager.update_informe_tecnico(informe_id, data):
                messagebox.showinfo("Éxito", "Cambios guardados correctamente.")
            else:
                messagebox.showerror("Error", "No se pudieron guardar los cambios.")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar cambios: {str(e)}")
            
    def save_informe_changes(self, informe_id):
        """Guarda los cambios realizados a un informe"""
        # Obtener datos del formulario
        data = {key: var.get() for key, var in self.edit_informe_vars.items()}
        
        # Guardar cambios
        if self.data_manager.update_informe_tecnico(informe_id, data):
            messagebox.showinfo("Éxito", "Cambios guardados correctamente.")
        else:
            messagebox.showerror("Error", "No se pudieron guardar los cambios.")
    
    def show_generate_word_window(self):
        """Muestra la ventana para generar documentos Word"""
        # Limpiar la ventana
        for widget in self.winfo_children():
            widget.destroy()
        
        # Crear frame principal
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        title = ctk.CTkLabel(main_frame, text="Generar Documento Word", 
                            font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=10)
        
        # Crear pestañas para los diferentes tipos de trabajo
        tab_view = ctk.CTkTabview(main_frame)
        tab_view.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Añadir pestañas
        tab_obra = tab_view.add("Obras en general")
        tab_informe = tab_view.add("Informes técnicos")
        
        # Configurar las pestañas
        self.setup_word_obra_tab(tab_obra)
        self.setup_word_informe_tab(tab_informe)
        
        # Botón para volver al menú principal
        btn_back = ctk.CTkButton(main_frame, text="Volver al Menú Principal", 
                                font=ctk.CTkFont(size=14),
                                command=self.create_main_menu)
        btn_back.pack(pady=10)
    
    def setup_word_obra_tab(self, parent):
        """Configura la pestaña para generar documentos Word de obras"""
        # Crear frame superior para búsqueda y selección
        selection_frame = ctk.CTkFrame(parent)
        selection_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Frame para búsqueda
        search_frame = ctk.CTkFrame(selection_frame)
        search_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Etiqueta para búsqueda
        ctk.CTkLabel(search_frame, text="Buscar por:").grid(row=0, column=0, padx=5, pady=5)
        
        # Opciones de búsqueda
        search_options = ["Profesional", "Comitente", "Partida Inmobiliaria", "Número GOP"]
        self.word_search_option_var = ctk.StringVar(value=search_options[0])
        option_menu = ctk.CTkOptionMenu(search_frame, values=search_options, variable=self.word_search_option_var)
        option_menu.grid(row=0, column=1, padx=5, pady=5)
        
        # Obtener listas de profesionales y comitentes existentes
        profesionales = self.data_manager.get_all_profesionales()
        comitentes = self.data_manager.get_all_comitentes()
        
        # Campo de búsqueda con autocompletado
        self.word_search_entry = AutocompleteEntry(search_frame, width=200, options=profesionales)
        self.word_search_entry.grid(row=0, column=2, padx=5, pady=5)
        
        # Actualizar opciones de autocompletado según el criterio seleccionado
        def update_search_options(*args):
            option = self.word_search_option_var.get()
            if option == "Profesional":
                self.word_search_entry.update_options(profesionales)
            elif option == "Comitente":
                self.word_search_entry.update_options(comitentes)
            else:
                self.word_search_entry.update_options([])  # Sin autocompletado para otros campos
        
        # Vincular el cambio de opción de búsqueda
        self.word_search_option_var.trace_add("write", update_search_options)
        
        # Botón de búsqueda
        search_button = ctk.CTkButton(search_frame, text="Buscar", command=self.search_word_obras)
        search_button.grid(row=0, column=3, padx=5, pady=5)
        
        # Configurar grid para ser responsive
        search_frame.grid_columnconfigure(2, weight=1)
        
        # Obtener todas las obras
        self.word_obras = self.data_manager.get_all_works("obra")
        
        # Crear lista de obras para mostrar
        obras_display = []
        for obra in self.word_obras:
            display = f"{obra['nombre_profesional']} - {obra['nombre_comitente']} ({obra['fecha']})"
            obras_display.append(display)
        
        # Selector de obra
        selection_menu_frame = ctk.CTkFrame(selection_frame)
        selection_menu_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ctk.CTkLabel(selection_menu_frame, text="Seleccionar Obra:").pack(side=tk.LEFT, padx=5)
        
        # Usamos una variable separada para el texto mostrado
        self.word_obra_selector_var = ctk.StringVar(value="Seleccione una obra" if obras_display else "No hay obras registradas")
        
        self.word_obra_selector = ctk.CTkOptionMenu(
            selection_menu_frame, 
            values=obras_display if obras_display else ["No hay obras registradas"],
            variable=self.word_obra_selector_var,
            command=self.on_word_obra_selected,
            dynamic_resizing=False,
            width=300
        )
        self.word_obra_selector.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Frame para las opciones de generación
        self.word_obra_frame = ctk.CTkFrame(parent)
        self.word_obra_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Indicar que seleccione una obra
        if not obras_display:
            ctk.CTkLabel(self.word_obra_frame, text="No hay obras disponibles", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
        else:
            # Seleccionar la primera obra
            self.word_obra_selector.set(obras_display[0])
            self.on_word_obra_selected(obras_display[0])

    def search_word_obras(self):
        """Busca obras para generar Word según el criterio seleccionado"""
        search_text = self.word_search_entry.get().strip().lower()
        search_option = self.word_search_option_var.get()
        
        if not search_text:
            messagebox.showwarning("Búsqueda vacía", "Por favor ingrese un texto para buscar")
            return
        
        # Obtener todas las obras de nuevo (por si se agregaron nuevas)
        self.word_obras = self.data_manager.get_all_works("obra")
        
        # Filtrar según el criterio
        filtered_obras = []
        
        for obra in self.word_obras:
            # Obtener obra completa para acceder a todos los campos
            obra_completa = self.data_manager.get_work_by_id("obra", obra["id"])
            
            if obra_completa:
                if search_option == "Profesional" and obra_completa["nombre_profesional"] and search_text in obra_completa["nombre_profesional"].lower():
                    filtered_obras.append(obra)
                elif search_option == "Comitente" and obra_completa["nombre_comitente"] and search_text in obra_completa["nombre_comitente"].lower():
                    filtered_obras.append(obra)
                elif search_option == "Partida Inmobiliaria" and obra_completa["nro_partida_inmobiliaria"] and search_text in str(obra_completa["nro_partida_inmobiliaria"]).lower():
                    filtered_obras.append(obra)
                elif search_option == "Número GOP" and obra_completa["nro_sistema_gop"] and search_text in str(obra_completa["nro_sistema_gop"]).lower():
                    filtered_obras.append(obra)
        
        # Actualizar la lista de obras mostradas
        if filtered_obras:
            obras_display = []
            for obra in filtered_obras:
                display = f"{obra['nombre_profesional']} - {obra['nombre_comitente']} ({obra['fecha']})"
                obras_display.append(display)
            
            # Guardar la lista filtrada
            self.word_obras = filtered_obras
            
            # Actualizar el selector
            self.word_obra_selector.configure(values=obras_display)
            self.word_obra_selector_var.set(obras_display[0])
            
            # Seleccionar la primera obra
            self.on_word_obra_selected(obras_display[0])
        else:
            messagebox.showinfo("Búsqueda", f"No se encontraron obras que coincidan con '{search_text}' en {search_option}")
    
    def on_word_obra_selected(self, selection):
        """Maneja la selección de una obra para generar documentos"""
        # Limpiar el frame de opciones
        for widget in self.word_obra_frame.winfo_children():
            widget.destroy()
        
        # Obtener el índice de la obra seleccionada
        try:
            if not selection or selection == "No hay obras registradas":
                ctk.CTkLabel(self.word_obra_frame, text="No hay obras disponibles", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
                return
                
            index = self.word_obra_selector.cget("values").index(selection)
            obra_id = self.word_obras[index]["id"]
            
            # Obtener los datos completos de la obra
            obra = self.data_manager.get_work_by_id("obra", obra_id)
            
            if obra:
                # Mostrar información de la obra
                info_frame = ctk.CTkFrame(self.word_obra_frame)
                info_frame.pack(fill=tk.X, padx=10, pady=10)
                
                info_text = f"Profesional: {obra['nombre_profesional']}\n"
                info_text += f"Comitente: {obra['nombre_comitente']}\n"
                info_text += f"Tipo: {obra['tipo_trabajo']}\n"
                info_text += f"Ubicación: {obra['ubicacion']}\n"
                info_text += f"Tasa de Sellado: {obra['tasa_sellado']}\n"
                info_text += f"Tasa de Visado: {obra['tasa_visado']}"
                
                ctk.CTkLabel(info_frame, text=info_text, justify=tk.LEFT).pack(pady=10, padx=10)
                
                # Crear botones para generar documentos
                buttons_frame = ctk.CTkFrame(self.word_obra_frame)
                buttons_frame.pack(fill=tk.X, padx=10, pady=10)
                
                # Botón para proforma de tasa de sellado
                btn_sellado = ctk.CTkButton(buttons_frame, text="Generar Proforma de Tasa de Sellado", 
                                        font=ctk.CTkFont(size=14),
                                        command=lambda: self.generate_obra_sellado(obra_id))
                btn_sellado.pack(pady=10, padx=10, fill=tk.X)
                
                # Botón para proforma de tasa de visado
                btn_visado = ctk.CTkButton(buttons_frame, text="Generar Proforma de Tasa de Visado", 
                                        font=ctk.CTkFont(size=14),
                                        command=lambda: self.generate_obra_visado(obra_id))
                btn_visado.pack(pady=10, padx=10, fill=tk.X)
            else:
                ctk.CTkLabel(self.word_obra_frame, text="No se pudo cargar la información de la obra", 
                            font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
        
        except Exception as e:
            print(f"Error al cargar la obra: {e}")
            ctk.CTkLabel(self.word_obra_frame, text=f"Error al cargar la obra: {str(e)}", 
                        font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
    
    def setup_word_informe_tab(self, parent):
        """Configura la pestaña para generar documentos Word de informes"""
        # Crear frame superior para búsqueda y selección
        selection_frame = ctk.CTkFrame(parent)
        selection_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Frame para búsqueda
        search_frame = ctk.CTkFrame(selection_frame)
        search_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Etiqueta para búsqueda
        ctk.CTkLabel(search_frame, text="Buscar por:").grid(row=0, column=0, padx=5, pady=5)
        
        # Opciones de búsqueda
        search_options = ["Profesional", "Comitente", "Tipo de Informe"]
        self.word_informe_search_option_var = ctk.StringVar(value=search_options[0])
        option_menu = ctk.CTkOptionMenu(search_frame, values=search_options, variable=self.word_informe_search_option_var)
        option_menu.grid(row=0, column=1, padx=5, pady=5)
        
        # Obtener listas de profesionales y comitentes existentes
        profesionales = self.data_manager.get_all_profesionales()
        comitentes = self.data_manager.get_all_comitentes()
        
        # Campo de búsqueda con autocompletado
        self.word_informe_search_entry = AutocompleteEntry(search_frame, width=200, options=profesionales)
        self.word_informe_search_entry.grid(row=0, column=2, padx=5, pady=5)
        
        # Actualizar opciones de autocompletado según el criterio seleccionado
        def update_search_options(*args):
            option = self.word_informe_search_option_var.get()
            if option == "Profesional":
                self.word_informe_search_entry.update_options(profesionales)
            elif option == "Comitente":
                self.word_informe_search_entry.update_options(comitentes)
            else:
                self.word_informe_search_entry.update_options([])  # Sin autocompletado para otros campos
        
        # Vincular el cambio de opción de búsqueda
        self.word_informe_search_option_var.trace_add("write", update_search_options)
        
        # Botón de búsqueda
        search_button = ctk.CTkButton(search_frame, text="Buscar", command=self.search_word_informes)
        search_button.grid(row=0, column=3, padx=5, pady=5)
        
        # Configurar grid para ser responsive
        search_frame.grid_columnconfigure(2, weight=1)
        
        # Obtener todos los informes
        self.word_informes = self.data_manager.get_all_works("informe")
        
        # Crear lista de informes para mostrar
        informes_display = []
        for informe in self.word_informes:
            display = f"{informe['profesional']} - {informe['comitente']} ({informe['fecha']})"
            informes_display.append(display)
        
        # Selector de informe
        selection_menu_frame = ctk.CTkFrame(selection_frame)
        selection_menu_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ctk.CTkLabel(selection_menu_frame, text="Seleccionar Informe:").pack(side=tk.LEFT, padx=5)
        
        # Usamos una variable separada para el texto mostrado
        self.word_informe_selector_var = ctk.StringVar(value="Seleccione un informe" if informes_display else "No hay informes registrados")
        
        self.word_informe_selector = ctk.CTkOptionMenu(
            selection_menu_frame, 
            values=informes_display if informes_display else ["No hay informes registrados"],
            variable=self.word_informe_selector_var,
            command=self.on_word_informe_selected,
            dynamic_resizing=False,
            width=300
        )
        self.word_informe_selector.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Frame para las opciones de generación
        self.word_informe_frame = ctk.CTkFrame(parent)
        self.word_informe_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Indicar que seleccione un informe
        if not informes_display:
            ctk.CTkLabel(self.word_informe_frame, text="No hay informes disponibles", 
                        font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
        else:
            # Seleccionar el primer informe
            self.word_informe_selector.set(informes_display[0])
            self.on_word_informe_selected(informes_display[0])

    def search_word_informes(self):
        """Busca informes para generar Word según el criterio seleccionado"""
        search_text = self.word_informe_search_entry.get().strip().lower()
        search_option = self.word_informe_search_option_var.get()
        
        if not search_text:
            messagebox.showwarning("Búsqueda vacía", "Por favor ingrese un texto para buscar")
            return
        
        # Obtener todos los informes de nuevo (por si se agregaron nuevos)
        self.word_informes = self.data_manager.get_all_works("informe")
        
        # Filtrar según el criterio
        filtered_informes = []
        
        for informe in self.word_informes:
            # Obtener informe completo para acceder a todos los campos
            informe_completo = self.data_manager.get_work_by_id("informe", informe["id"])
            
            if informe_completo:
                if search_option == "Profesional" and informe_completo["profesional"] and search_text in informe_completo["profesional"].lower():
                    filtered_informes.append(informe)
                elif search_option == "Comitente" and informe_completo["comitente"] and search_text in informe_completo["comitente"].lower():
                    filtered_informes.append(informe)
                elif search_option == "Tipo de trabajo" and informe_completo["tipo_trabajo"] and search_text in str(informe_completo["tipo_trabajo"]).lower():
                    filtered_informes.append(informe)
        
        # Actualizar la lista de informes mostrados
        if filtered_informes:
            informes_display = []
            for informe in filtered_informes:
                display = f"{informe['profesional']} - {informe['comitente']} ({informe['fecha']})"
                informes_display.append(display)
            
            # Guardar la lista filtrada
            self.word_informes = filtered_informes
            
            # Actualizar el selector
            self.word_informe_selector.configure(values=informes_display)
            self.word_informe_selector_var.set(informes_display[0])
            
            # Seleccionar el primer informe
            self.on_word_informe_selected(informes_display[0])
        else:
            messagebox.showinfo("Búsqueda", f"No se encontraron informes que coincidan con '{search_text}' en {search_option}")
            
    def on_word_informe_selected(self, selection):
        """Maneja la selección de un informe para generar documentos"""
        # Limpiar el frame de opciones
        for widget in self.word_informe_frame.winfo_children():
            widget.destroy()
        
        # Obtener el índice del informe seleccionado
        try:
            if not selection or selection == "No hay informes registrados":
                ctk.CTkLabel(self.word_informe_frame, text="No hay informes disponibles", 
                            font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
                return
                
            index = self.word_informe_selector.cget("values").index(selection)
            informe_id = self.word_informes[index]["id"]
            
            # Obtener los datos completos del informe
            informe = self.data_manager.get_work_by_id("informe", informe_id)
            
            if informe:
                # Mostrar información del informe
                info_frame = ctk.CTkFrame(self.word_informe_frame)
                info_frame.pack(fill=tk.X, padx=10, pady=10)
                
                info_text = f"Profesional: {informe['profesional']}\n"
                info_text += f"Comitente: {informe['comitente']}\n"
                info_text += f"Tipo: {informe['tipo_trabajo']}\n"
                info_text += f"Detalle: {informe['detalle']}\n"
                info_text += f"Tasa de Sellado: {informe['tasa_sellado']}"
                
                ctk.CTkLabel(info_frame, text=info_text, justify=tk.LEFT).pack(pady=10, padx=10)
                
                # Crear botón para generar documento
                buttons_frame = ctk.CTkFrame(self.word_informe_frame)
                buttons_frame.pack(fill=tk.X, padx=10, pady=10)
                
                # Botón para generar proforma de informe técnico
                btn_informe = ctk.CTkButton(buttons_frame, text="Generar Proforma de Informe Técnico", 
                                        font=ctk.CTkFont(size=14),
                                        command=lambda: self.generate_informe(informe_id))
                btn_informe.pack(pady=10, padx=10, fill=tk.X)
            else:
                ctk.CTkLabel(self.word_informe_frame, text="No se pudo cargar la información del informe", 
                            font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
        
        except Exception as e:
            print(f"Error al cargar el informe: {e}")
            ctk.CTkLabel(self.word_informe_frame, text=f"Error al cargar el informe: {str(e)}", 
                        font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
    
    def generate_obra_sellado(self, obra_id):
        """Genera un documento Word para la tasa de sellado de una obra"""
        # Obtener los datos de la obra
        obra = self.data_manager.get_work_by_id("obra", obra_id)
        
        if obra:
            # Determinar la ruta de salida
            if obra["formato"] == "Digital" and obra["ruta_carpeta"]:
                output_dir = Path(obra["ruta_carpeta"])
            else:
                output_dir = Path.cwd()
            
            output_path = output_dir / f"Sellado_{obra['nombre_comitente']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            
            # Generar el documento
            result = self.word_generator.generate_obra_sellado(obra, output_path)
            
            if result.startswith("Error") or result.startswith("Faltan"):
                messagebox.showerror("Error", result)
            else:
                messagebox.showinfo("Éxito", f"Documento generado correctamente en:\n{result}")
                
                # Abrir la carpeta donde se guardó el documento
                self.file_manager.open_folder(output_dir)
                
    def generate_obra_visado(self, obra_id):
        """Genera un documento Word para la tasa de visado de una obra"""
        # Obtener los datos de la obra
        obra = self.data_manager.get_work_by_id("obra", obra_id)
        
        if obra:
            # Determinar la ruta de salida
            if obra["formato"] == "Digital" and obra["ruta_carpeta"]:
                output_dir = Path(obra["ruta_carpeta"])
            else:
                output_dir = Path.cwd()
            
            output_path = output_dir / f"Visado_{obra['nombre_comitente']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            
            # Generar el documento
            result = self.word_generator.generate_obra_visado(obra, output_path)
            
            if result.startswith("Error") or result.startswith("Faltan"):
                messagebox.showerror("Error", result)
            else:
                messagebox.showinfo("Éxito", f"Documento generado correctamente en:\n{result}")
                
                # Asegurarse de abrir la carpeta donde se guardó el documento
                self.file_manager.open_folder(output_dir)
        else:
            messagebox.showerror("Error", "No se encontró la información de la obra")