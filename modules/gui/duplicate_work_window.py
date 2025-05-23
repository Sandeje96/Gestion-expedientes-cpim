import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox
from .autocomplete_widget import AutocompleteEntry


class DuplicateWorkWindow:
    def __init__(self, parent, data_manager, file_manager, return_callback, new_record_callback):
        self.parent = parent
        self.data_manager = data_manager
        self.file_manager = file_manager
        self.return_callback = return_callback
        self.new_record_callback = new_record_callback
        
        self.setup_window()
    
    def setup_window(self):
        """Configura la ventana de duplicar trabajo"""
        # Crear frame principal
        main_frame = ctk.CTkFrame(self.parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        title = ctk.CTkLabel(
            main_frame, 
            text="Repetir Trabajo con Otro Profesional", 
            font=ctk.CTkFont(size=20, weight="bold")
        )
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
        btn_back = ctk.CTkButton(
            main_frame, 
            text="Volver al Menú Principal", 
            font=ctk.CTkFont(size=14),
            command=self.return_callback
        )
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
                
                btn_duplicate = ctk.CTkButton(
                    btn_frame, 
                    text="Duplicar con Nuevo Profesional", 
                    font=ctk.CTkFont(size=14),
                    command=lambda: self.repeat_obra_with_new_professional(obra_id)
                )
                btn_duplicate.pack(pady=10, padx=10, fill=tk.X)
                
                # Instrucciones
                instructions = "Al hacer clic en 'Duplicar con Nuevo Profesional', se abrirá el formulario de nuevo registro con todos los datos de este trabajo pre-cargados, excepto los datos del profesional.\n\nSolo tendrá que completar la información del nuevo profesional para registrar el trabajo."
                ctk.CTkLabel(self.duplicate_obra_frame, text=instructions, wraplength=600, justify=tk.LEFT).pack(pady=20, padx=10)
            
            else:
                ctk.CTkLabel(self.duplicate_obra_frame, text="No se pudo cargar la información de la obra", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
        
        except Exception as e:
            print(f"Error al cargar la obra para duplicar: {e}")
            ctk.CTkLabel(self.duplicate_obra_frame, text=f"Error al cargar la obra: {str(e)}", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
    
    def repeat_obra_with_new_professional(self, obra_id):
        """Prepara un nuevo registro basado en una obra existente pero para otro profesional"""
        # Obtener los datos completos de la obra
        obra = self.data_manager.get_work_by_id("obra", obra_id)
        
        if not obra:
            messagebox.showerror("Error", "No se pudo cargar la información de la obra original")
            return
        
        # Mostrar mensaje informativo
        messagebox.showinfo("Repetir trabajo", 
                        "Se han copiado los datos del trabajo seleccionado.\n\n"
                        "Por favor, complete la información del nuevo profesional y guarde el registro.")
        
        # Cambiar a la ventana de nuevo registro y poblar con los datos
        self.new_record_callback()
        
        # Necesitamos una forma de pasar los datos a la nueva ventana
        # Por ahora, guardaremos los datos temporalmente en el parent
        if hasattr(self.parent, 'current_window') and hasattr(self.parent.current_window, 'populate_obra_fields'):
            self.parent.current_window.populate_obra_fields(obra)