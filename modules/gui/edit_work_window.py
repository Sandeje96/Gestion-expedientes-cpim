import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox
from .autocomplete_widget import AutocompleteEntry


class EditWorkWindow:
    def __init__(self, parent, data_manager, file_manager, return_callback):
        self.parent = parent
        self.data_manager = data_manager
        self.file_manager = file_manager
        self.return_callback = return_callback
        
        self.setup_window()
    
    def setup_window(self):
        """Configura la ventana de editar trabajo"""
        # Crear frame principal
        main_frame = ctk.CTkFrame(self.parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        title = ctk.CTkLabel(
            main_frame, 
            text="Editar Trabajo", 
            font=ctk.CTkFont(size=20, weight="bold")
        )
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
        btn_back = ctk.CTkButton(
            main_frame, 
            text="Volver al Menú Principal", 
            font=ctk.CTkFont(size=14),
            command=self.return_callback
        )
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
        self.edit_obra_vars = {}
        
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
                btn_repeat = ctk.CTkButton(
                    info_frame, 
                    text="Repetir con otro Profesional", 
                    command=lambda: self.repeat_obra_with_new_professional(obra_id)
                )
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
                
                # Visados específicos
                visados = [
                    ("Visado de instalación de Gas:", "visado_gas"),
                    ("Visado de instalación de Salubridad:", "visado_salubridad"),
                    ("Visado de instalación eléctrica:", "visado_electrica"),
                    ("Visado de instalación electromecánica:", "visado_electromecanica")
                ]
                
                for label_text, field_name in visados:
                    ctk.CTkLabel(edit_frame, text=label_text).grid(row=row, column=0, sticky="w", pady=5, padx=5)
                    entry = ctk.CTkEntry(edit_frame)
                    entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                    entry.insert(0, obra[field_name] if obra[field_name] else "")
                    self.edit_obra_vars[field_name] = entry
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
                
                # Campos adicionales
                additional_fields = [
                    ("Nro. Expediente CPIM:", "nro_expediente_cpim"),
                    ("Fecha de Salida:", "fecha_salida"),
                    ("Persona que Retira:", "persona_retira"),
                    ("Nro. de Caja:", "nro_caja")
                ]
                
                for label_text, field_name in additional_fields:
                    ctk.CTkLabel(edit_frame, text=label_text).grid(row=row, column=0, sticky="w", pady=5, padx=5)
                    entry = ctk.CTkEntry(edit_frame)
                    entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                    entry.insert(0, obra[field_name] if obra[field_name] else "")
                    self.edit_obra_vars[field_name] = entry
                    row += 1
                
                # Si es formato digital, mostrar botón para abrir carpeta
                if obra["formato"] == "Digital" and obra["ruta_carpeta"]:
                    btn_open_folder = ctk.CTkButton(
                        edit_frame, 
                        text="Abrir Carpeta", 
                        command=lambda: self.file_manager.open_folder(obra["ruta_carpeta"])
                    )
                    btn_open_folder.grid(row=row, column=0, pady=10, padx=5)
                    row += 1
                
                # Configurar el grid para que sea responsive
                edit_frame.grid_columnconfigure(1, weight=1)
                
                # Botón para guardar cambios
                btn_save = ctk.CTkButton(
                    edit_frame, 
                    text="Guardar Cambios", 
                    font=ctk.CTkFont(size=14),
                    command=lambda: self.save_obra_changes(obra_id)
                )
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
        
        # Aquí necesitaríamos acceso a la ventana de nuevo registro
        # Por ahora mostramos un mensaje
        messagebox.showinfo("Repetir trabajo", 
                        "Para repetir este trabajo con otro profesional, vaya al menú principal y seleccione 'Repetir Trabajo con Otro Profesional'.")
    
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
        
        # Variables para los campos editables
        self.edit_informe_vars = {}
        
        # Indicar que no hay informe seleccionado
        ctk.CTkLabel(self.informe_edit_frame, text="Seleccione un informe para editar", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
        
        # Seleccionar el primer informe si hay alguno
        if informes_display:
            self.informe_selector.set(informes_display[0])
            self.on_informe_selected(informes_display[0])
    
    def search_informes(self):
        """Busca informes según el criterio seleccionado"""
        search_text = self.informe_search_entry.get().strip().lower()
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
                
                # Campos adicionales
                additional_fields = [
                    ("Nro. Expediente CPIM:", "nro_expediente_cpim"),
                    ("Fecha de Salida:", "fecha_salida"),
                    ("Persona que Retira:", "persona_retira"),
                    ("Nro. de Caja:", "nro_caja")
                ]
                
                for label_text, field_name in additional_fields:
                    ctk.CTkLabel(edit_frame, text=label_text).grid(row=row, column=0, sticky="w", pady=5, padx=5)
                    entry = ctk.CTkEntry(edit_frame)
                    entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)
                    entry.insert(0, informe[field_name] if informe[field_name] else "")
                    self.edit_informe_vars[field_name] = entry
                    row += 1
                
                # Si es formato digital, mostrar botón para abrir carpeta
                if informe["formato"] == "Digital" and informe["ruta_carpeta"]:
                    btn_open_folder = ctk.CTkButton(
                        edit_frame, 
                        text="Abrir Carpeta", 
                        command=lambda: self.file_manager.open_folder(informe["ruta_carpeta"])
                    )
                    btn_open_folder.grid(row=row, column=0, pady=10, padx=5)
                    row += 1
                    
                # Configurar el grid para que sea responsive
                edit_frame.grid_columnconfigure(1, weight=1)
                
                # Botón para guardar cambios
                btn_save = ctk.CTkButton(
                    edit_frame, 
                    text="Guardar Cambios", 
                    font=ctk.CTkFont(size=14),
                    command=lambda: self.save_informe_changes(informe_id)
                )
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