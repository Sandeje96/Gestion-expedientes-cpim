import os
import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox, filedialog
import shutil
from datetime import datetime
from config import PROFESIONES, TIPOS_OBRA, TIPOS_INFORME
from .autocomplete_widget import AutocompleteEntry
from .currency_entry import CurrencyEntry


class NewRecordWindow:
    def __init__(self, parent, data_manager, file_manager, return_callback):
        self.parent = parent
        self.data_manager = data_manager
        self.file_manager = file_manager
        self.return_callback = return_callback
        
        self.setup_window()
    
    def setup_window(self):
        """Configura la ventana de nuevo registro"""
        # Crear frame principal
        main_frame = ctk.CTkFrame(self.parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        title = ctk.CTkLabel(
            main_frame, 
            text="Nuevo Registro", 
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title.pack(pady=10)
        
        # Crear pestañas para los diferentes tipos de trabajo
        self.tab_view = ctk.CTkTabview(main_frame)
        self.tab_view.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Añadir pestañas
        tab_obra = self.tab_view.add("Obra en general")
        tab_informe = self.tab_view.add("Informe técnico")
        
        # Contenido de las pestañas
        self.setup_obra_tab(tab_obra)
        self.setup_informe_tab(tab_informe)
        
        # Botón para volver al menú principal
        btn_back = ctk.CTkButton(
            main_frame, 
            text="Volver al Menú Principal", 
            font=ctk.CTkFont(size=14),
            command=self.return_callback
        )
        btn_back.pack(pady=10)
    
    def validate_length(self, text, max_length):
        """Valida que el texto no exceda la longitud máxima"""
        return len(text) <= max_length
    
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
            "visado_electromecanica": ctk.StringVar(),
            "whatsapp_profesional": ctk.StringVar(),
            "whatsapp_tramitador": ctk.StringVar()
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
        
        # Configurar el frame de archivos
        self.setup_file_upload_section(self.upload_files_frame)
        row += 1
        
        # Ajustar visibilidad inicial de campos según formato
        self.toggle_formato_fields()
        
        # Tipo de trabajo
        ctk.CTkLabel(scroll_frame, text="Tipo de trabajo:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        ctk.CTkOptionMenu(scroll_frame, values=TIPOS_OBRA, variable=self.obra_vars["tipo_trabajo"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Nombre del profesional (con autocompletado y límite de caracteres)
        ctk.CTkLabel(scroll_frame, text="Nombre del Profesional:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        self.prof_entry = AutocompleteEntry(scroll_frame, options=profesionales)
        # Configurar callback para autocompletar WhatsApp cuando cambie el profesional
        self.prof_entry.entry_var.trace_add("write", self.on_profesional_change_obra)
        
        # APLICAR LÍMITE DE 50 CARACTERES AL PROFESIONAL
        validate_prof_cmd = (self.parent.register(lambda text: self.validate_length(text, 50)), '%P')
        self.prof_entry.entry.configure(validate='key', validatecommand=validate_prof_cmd)
        
        self.prof_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        self.obra_vars["nombre_profesional"] = self.prof_entry
        
        # Label informativo para profesional
        ctk.CTkLabel(scroll_frame, text="(máx. 50 caracteres)", font=("Arial", 8), text_color="gray").grid(row=row, column=2, sticky="w", padx=5)
        row += 1
        
        # Nombre del comitente (con autocompletado y límite de caracteres)
        ctk.CTkLabel(scroll_frame, text="Nombre del Comitente:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        self.comit_entry = AutocompleteEntry(scroll_frame, options=comitentes)
        
        # APLICAR LÍMITE DE 80 CARACTERES AL COMITENTE
        validate_comit_cmd = (self.parent.register(lambda text: self.validate_length(text, 50)), '%P')
        self.comit_entry.entry.configure(validate='key', validatecommand=validate_comit_cmd)
        
        self.comit_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        self.obra_vars["nombre_comitente"] = self.comit_entry
        
        # Label informativo para comitente
        ctk.CTkLabel(scroll_frame, text="(usa - para separar muchos comitentes)", font=("Arial", 8), text_color="gray").grid(row=row, column=2, sticky="w", padx=5)
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
        CurrencyEntry(scroll_frame, textvariable=self.obra_vars["tasa_sellado"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1

        # Tasa de visado (mantenemos este campo para compatibilidad)
        ctk.CTkLabel(scroll_frame, text="Tasa de Visado (General):").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        CurrencyEntry(scroll_frame, textvariable=self.obra_vars["tasa_visado"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1

        # Añadir un separador o título para la sección de visados específicos
        ctk.CTkLabel(scroll_frame, text="Visados Específicos", font=("Arial", 12, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", pady=(15, 5), padx=5)
        row += 1

        # Visados específicos
        visados = [
            ("Visado de instalación de Gas:", "visado_gas"),
            ("Visado de instalación de Salubridad:", "visado_salubridad"),
            ("Visado de instalación eléctrica:", "visado_electrica"),
            ("Visado de instalación electromecánica:", "visado_electromecanica")
        ]

        for label_text, var_name in visados:
            ctk.CTkLabel(scroll_frame, text=label_text).grid(row=row, column=0, sticky="w", pady=5, padx=5)
            CurrencyEntry(scroll_frame, textvariable=self.obra_vars[var_name]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
            row += 1
        
        # Añadir separador para WhatsApp
        ctk.CTkLabel(scroll_frame, text="Datos de Contacto", font=("Arial", 12, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", pady=(15, 5), padx=5)
        row += 1
        
        # WhatsApp del profesional
        ctk.CTkLabel(scroll_frame, text="WhatsApp Profesional:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        whatsapp_prof_entry = ctk.CTkEntry(scroll_frame, textvariable=self.obra_vars["whatsapp_profesional"], placeholder_text="Ej: +5493755123456 o 3755123456")
        whatsapp_prof_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # WhatsApp del tramitador
        ctk.CTkLabel(scroll_frame, text="WhatsApp Tramitador (opcional):").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        whatsapp_tram_entry = ctk.CTkEntry(scroll_frame, textvariable=self.obra_vars["whatsapp_tramitador"], placeholder_text="Ej: +5493755123456 o 3755123456")
        whatsapp_tram_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Configurar el grid para que sea responsive
        scroll_frame.grid_columnconfigure(1, weight=1)
        
        # Botón para guardar
        btn_save = ctk.CTkButton(
            scroll_frame, 
            text="Guardar Obra", 
            font=ctk.CTkFont(size=14),
            command=self.save_obra
        )
        btn_save.grid(row=row, column=0, columnspan=2, pady=20, padx=5, sticky="ew")
    
    def setup_file_upload_section(self, parent_frame):
        """Configura la sección de subida de archivos"""
        # Etiqueta para archivos
        ctk.CTkLabel(parent_frame, text="Archivos seleccionados:").pack(anchor="w", padx=5, pady=5)
        
        # Lista de archivos seleccionados
        self.files_listbox = tk.Listbox(parent_frame, height=5, width=50)
        self.files_listbox.pack(fill="x", padx=5, pady=5)
        
        # Botones para archivos
        files_buttons_frame = ctk.CTkFrame(parent_frame)
        files_buttons_frame.pack(fill="x", padx=5, pady=5)
        
        # Botón para añadir archivos
        add_files_btn = ctk.CTkButton(
            files_buttons_frame, 
            text="Añadir Archivos", 
            command=self.add_files_to_list
        )
        add_files_btn.pack(side="left", padx=5)
        
        # Botón para eliminar archivos seleccionados
        remove_files_btn = ctk.CTkButton(
            files_buttons_frame, 
            text="Eliminar Seleccionado", 
            command=self.remove_selected_file
        )
        remove_files_btn.pack(side="left", padx=5)
        
        # Guardar lista de archivos seleccionados
        self.selected_files = []
    
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
            "tasa_sellado": ctk.StringVar(),
            "whatsapp_profesional": ctk.StringVar(),
            "whatsapp_tramitador": ctk.StringVar()
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
        
        # Configurar el frame de archivos para informes
        self.setup_informe_file_upload_section(self.informe_upload_files_frame)
        row += 1
        
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
        
        # Profesional (con autocompletado y límite de caracteres)
        ctk.CTkLabel(scroll_frame, text="Profesional:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        self.informe_prof_entry = AutocompleteEntry(scroll_frame, options=profesionales)
        # Configurar callback para autocompletar WhatsApp cuando cambie el profesional
        self.informe_prof_entry.entry_var.trace_add("write", self.on_profesional_change_informe)
        
        # APLICAR LÍMITE DE 50 CARACTERES AL PROFESIONAL
        validate_inf_prof_cmd = (self.parent.register(lambda text: self.validate_length(text, 50)), '%P')
        self.informe_prof_entry.entry.configure(validate='key', validatecommand=validate_inf_prof_cmd)
        
        self.informe_prof_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        self.informe_vars["profesional"] = self.informe_prof_entry
        
        # Label informativo para profesional
        ctk.CTkLabel(scroll_frame, text="(máx. 50 caracteres)", font=("Arial", 8), text_color="gray").grid(row=row, column=2, sticky="w", padx=5)
        row += 1
        
        # Comitente (con autocompletado y límite de caracteres)
        ctk.CTkLabel(scroll_frame, text="Comitente:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        self.informe_comit_entry = AutocompleteEntry(scroll_frame, options=comitentes)
        
        # APLICAR LÍMITE DE 80 CARACTERES AL COMITENTE
        validate_inf_comit_cmd = (self.parent.register(lambda text: self.validate_length(text, 50)), '%P')
        self.informe_comit_entry.entry.configure(validate='key', validatecommand=validate_inf_comit_cmd)
        
        self.informe_comit_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        self.informe_vars["comitente"] = self.informe_comit_entry
        
        # Label informativo para comitente
        ctk.CTkLabel(scroll_frame, text="(usa - para separar muchos comitentes)", font=("Arial", 8), text_color="gray").grid(row=row, column=2, sticky="w", padx=5)
        row += 1
        
        # Tasa de sellado
        ctk.CTkLabel(scroll_frame, text="Tasa de Sellado:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        CurrencyEntry(scroll_frame, textvariable=self.informe_vars["tasa_sellado"]).grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Añadir separador para WhatsApp
        ctk.CTkLabel(scroll_frame, text="Datos de Contacto", font=("Arial", 12, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", pady=(15, 5), padx=5)
        row += 1
        
        # WhatsApp del profesional
        ctk.CTkLabel(scroll_frame, text="WhatsApp Profesional:").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        whatsapp_prof_entry = ctk.CTkEntry(scroll_frame, textvariable=self.informe_vars["whatsapp_profesional"], placeholder_text="Ej: +5493755123456 o 3755123456")
        whatsapp_prof_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # WhatsApp del tramitador
        ctk.CTkLabel(scroll_frame, text="WhatsApp Tramitador (opcional):").grid(row=row, column=0, sticky="w", pady=5, padx=5)
        whatsapp_tram_entry = ctk.CTkEntry(scroll_frame, textvariable=self.informe_vars["whatsapp_tramitador"], placeholder_text="Ej: +5493755123456 o 3755123456")
        whatsapp_tram_entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
        row += 1
        
        # Configurar el grid para que sea responsive
        scroll_frame.grid_columnconfigure(1, weight=1)
        
        # Botón para guardar
        btn_save = ctk.CTkButton(
            scroll_frame, 
            text="Guardar Informe", 
            font=ctk.CTkFont(size=14),
            command=self.save_informe
        )
        btn_save.grid(row=row, column=0, columnspan=2, pady=20, padx=5, sticky="ew")
    
    def setup_informe_file_upload_section(self, parent_frame):
        """Configura la sección de subida de archivos para informes"""
        # Etiqueta para archivos
        ctk.CTkLabel(parent_frame, text="Archivos seleccionados:").pack(anchor="w", padx=5, pady=5)
        
        # Lista de archivos seleccionados
        self.informe_files_listbox = tk.Listbox(parent_frame, height=5, width=50)
        self.informe_files_listbox.pack(fill="x", padx=5, pady=5)
        
        # Botones para archivos
        files_buttons_frame = ctk.CTkFrame(parent_frame)
        files_buttons_frame.pack(fill="x", padx=5, pady=5)
        
        # Botón para añadir archivos
        add_files_btn = ctk.CTkButton(
            files_buttons_frame, 
            text="Añadir Archivos", 
            command=self.add_files_to_informe_list
        )
        add_files_btn.pack(side="left", padx=5)
        
        # Botón para eliminar archivos seleccionados
        remove_files_btn = ctk.CTkButton(
            files_buttons_frame, 
            text="Eliminar Seleccionado", 
            command=self.remove_selected_informe_file
        )
        remove_files_btn.pack(side="left", padx=5)
        
        # Guardar lista de archivos seleccionados
        self.informe_selected_files = []
    
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
        
        # Obtener datos del formulario y limpiar espacios en blanco
        data = {}
        for key, var in self.obra_vars.items():
            if key in ["nombre_profesional", "nombre_comitente"]:
                # Estos son widgets AutocompleteEntry - limpiar espacios
                data[key] = var.get().strip()
            else:
                # Estos son StringVar - limpiar espacios
                data[key] = var.get().strip()
        
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
        
        # Obtener datos del formulario y limpiar espacios en blanco
        data = {}
        for key, var in self.informe_vars.items():
            if key in ["profesional", "comitente"]:
                # Estos son widgets AutocompleteEntry - limpiar espacios
                data[key] = var.get().strip()
            else:
                # Estos son StringVar - limpiar espacios
                data[key] = var.get().strip()
        
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
    
    def populate_obra_fields(self, obra_data):
        """Rellena los campos del formulario de obra con datos existentes (para duplicar)"""
        try:
            # Formato
            self.formato_var.set(obra_data["formato"])
            self.toggle_formato_fields()
            
            # Números de copias o GOP según formato
            if obra_data["formato"] == "Físico":
                self.obra_vars["nro_copias"].set(obra_data["nro_copias"])
            else:
                self.obra_vars["nro_sistema_gop"].set(obra_data["nro_sistema_gop"])
            
            # Tipo de trabajo
            self.obra_vars["tipo_trabajo"].set(obra_data["tipo_trabajo"])
            
            # Datos del comitente y ubicación
            self.obra_vars["nombre_comitente"].set(obra_data["nombre_comitente"])
            # Forzar que se oculte el dropdown de autocompletado
            self.comit_entry.hide_dropdown()
            
            self.obra_vars["ubicacion"].set(obra_data["ubicacion"])
            self.obra_vars["nro_expte_municipal"].set(obra_data["nro_expte_municipal"])
            self.obra_vars["nro_partida_inmobiliaria"].set(obra_data["nro_partida_inmobiliaria"])
            
            # Tasas
            self.obra_vars["tasa_sellado"].set(obra_data["tasa_sellado"])
            self.obra_vars["tasa_visado"].set(obra_data["tasa_visado"])
            
            # Nuevos campos de visado
            self.obra_vars["visado_gas"].set(obra_data["visado_gas"])
            self.obra_vars["visado_salubridad"].set(obra_data["visado_salubridad"])
            self.obra_vars["visado_electrica"].set(obra_data["visado_electrica"])
            self.obra_vars["visado_electromecanica"].set(obra_data["visado_electromecanica"])
            
            # Campos de WhatsApp (si existen en los datos)
            self.obra_vars["whatsapp_profesional"].set(obra_data.get("whatsapp_profesional", ""))
            self.obra_vars["whatsapp_tramitador"].set(obra_data.get("whatsapp_tramitador", ""))
            
            # Limpiar específicamente los datos del profesional
            self.obra_vars["profesion"].set("")
            self.obra_vars["nombre_profesional"].set("")
            
            # Cambiar a la pestaña de obra en general
            self.tab_view.set("Obra en general")
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron copiar todos los datos: {str(e)}")

    def on_profesional_change_obra(self, *args):
        """Se ejecuta cuando cambia el nombre del profesional en obras"""
        try:
            nombre_profesional = self.prof_entry.get().strip()
            if nombre_profesional:
                # Buscar si este profesional ya tiene un WhatsApp registrado
                whatsapp = self.data_manager.get_whatsapp_by_profesional(nombre_profesional)
                if whatsapp:
                    # Autocompletar el campo de WhatsApp
                    self.obra_vars["whatsapp_profesional"].set(whatsapp)
        except Exception as e:
            print(f"Error al autocompletar WhatsApp en obra: {e}")

    def on_profesional_change_informe(self, *args):
        """Se ejecuta cuando cambia el nombre del profesional en informes"""
        try:
            nombre_profesional = self.informe_prof_entry.get().strip()
            if nombre_profesional:
                # Buscar si este profesional ya tiene un WhatsApp registrado
                whatsapp = self.data_manager.get_whatsapp_by_profesional(nombre_profesional)
                if whatsapp:
                    # Autocompletar el campo de WhatsApp
                    self.informe_vars["whatsapp_profesional"].set(whatsapp)
        except Exception as e:
            print(f"Error al autocompletar WhatsApp en informe: {e}")