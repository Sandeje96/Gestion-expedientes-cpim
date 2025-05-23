import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime
from pathlib import Path
import subprocess
import os
from .autocomplete_widget import AutocompleteEntry


class GenerateWordWindow:
    def __init__(self, parent, data_manager, file_manager, word_generator, return_callback):
        self.parent = parent
        self.data_manager = data_manager
        self.file_manager = file_manager
        self.word_generator = word_generator
        self.return_callback = return_callback
        
        self.setup_window()
    
    def setup_window(self):
        """Configura la ventana de generar documentos Word"""
        # Crear frame principal
        main_frame = ctk.CTkFrame(self.parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        title = ctk.CTkLabel(
            main_frame, 
            text="Generar Documento Word", 
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
        self.setup_word_obra_tab(tab_obra)
        self.setup_word_informe_tab(tab_informe)
        
        # Botón para volver al menú principal
        btn_back = ctk.CTkButton(
            main_frame, 
            text="Volver al Menú Principal", 
            font=ctk.CTkFont(size=14),
            command=self.return_callback
        )
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
                btn_sellado = ctk.CTkButton(
                    buttons_frame, 
                    text="Generar Proforma de Tasa de Sellado", 
                    font=ctk.CTkFont(size=14),
                    command=lambda: self.generate_obra_sellado(obra_id)
                )
                btn_sellado.pack(pady=10, padx=10, fill=tk.X)
                
                # Botón para proforma de tasa de visado
                btn_visado = ctk.CTkButton(
                    buttons_frame, 
                    text="Generar Proforma de Tasa de Visado", 
                    font=ctk.CTkFont(size=14),
                    command=lambda: self.generate_obra_visado(obra_id)
                )
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
        
        # TODO: Completar la configuración de la pestaña de informes
        # Por ahora, mostrar un mensaje temporal
        temp_label = ctk.CTkLabel(parent, text="Funcionalidad de informes en desarrollo", 
                                 font=ctk.CTkFont(size=16, weight="bold"))
        temp_label.pack(pady=50)
    
    def generate_obra_sellado(self, obra_id):
        """Genera el documento Word de proforma de tasa de sellado"""
        try:
            # Obtener los datos de la obra
            obra = self.data_manager.get_work_by_id("obra", obra_id)
            
            if not obra:
                messagebox.showerror("Error", "No se pudo obtener la información de la obra")
                return
            
            # Determinar la carpeta de destino correcta
            if obra.get("ruta_carpeta"):
                # Si tiene ruta de carpeta, usar esa
                output_dir = Path(obra["ruta_carpeta"])
            else:
                # Si no tiene ruta (trabajos físicos), crear la estructura de carpetas
                output_dir = self.file_manager.create_folder_structure(obra, "obra")
            
            # Crear el nombre del archivo
            filename = f"Proforma_Sellado_{obra['nombre_comitente']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            output_path = output_dir / filename
            
            # Generar el documento
            result = self.word_generator.generate_obra_sellado(obra, output_path)
            
            if result.startswith("Error") or result.startswith("Faltan"):
                messagebox.showerror("Error", result)
            else:
                messagebox.showinfo("Éxito", f"Documento generado correctamente:\n{filename}")
                
                # Abrir el archivo Word generado
                import subprocess
                import os
                
                if os.name == 'nt':  # Windows
                    os.startfile(str(output_path))
                elif os.name == 'posix':  # macOS y Linux
                    if 'darwin' in os.sys.platform:
                        subprocess.call(['open', str(output_path)])
                    else:
                        subprocess.call(['xdg-open', str(output_path)])
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar documento: {str(e)}")
    
    def generate_obra_visado(self, obra_id):
        """Genera el documento Word de proforma de tasa de visado"""
        try:
            # Obtener los datos de la obra
            obra = self.data_manager.get_work_by_id("obra", obra_id)
            
            if not obra:
                messagebox.showerror("Error", "No se pudo obtener la información de la obra")
                return
            
            # Determinar la carpeta de destino correcta
            if obra.get("ruta_carpeta"):
                # Si tiene ruta de carpeta, usar esa
                output_dir = Path(obra["ruta_carpeta"])
            else:
                # Si no tiene ruta (trabajos físicos), crear la estructura de carpetas
                output_dir = self.file_manager.create_folder_structure(obra, "obra")
            
            # Crear el nombre del archivo
            filename = f"Proforma_Visado_{obra['nombre_comitente']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            output_path = output_dir / filename
            
            # Generar el documento
            result = self.word_generator.generate_obra_visado(obra, output_path)
            
            if result.startswith("Error") or result.startswith("Faltan"):
                messagebox.showerror("Error", result)
            else:
                messagebox.showinfo("Éxito", f"Documento generado correctamente:\n{filename}")
                
                # Abrir el archivo Word generado
                import subprocess
                import os
                
                if os.name == 'nt':  # Windows
                    os.startfile(str(output_path))
                elif os.name == 'posix':  # macOS y Linux
                    if 'darwin' in os.sys.platform:
                        subprocess.call(['open', str(output_path)])
                    else:
                        subprocess.call(['xdg-open', str(output_path)])
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar documento: {str(e)}")