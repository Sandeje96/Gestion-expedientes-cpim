import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox, filedialog
from datetime import datetime, timedelta
import calendar
from pathlib import Path
from ..tasas_analyzer import TasasAnalyzer
from config import BASE_PATH


class TasasAnalysisWindow:
    def __init__(self, parent, data_manager, return_callback):
        self.parent = parent
        self.data_manager = data_manager
        self.return_callback = return_callback
        self.tasas_analyzer = TasasAnalyzer(data_manager)
        self.current_analysis = None
        
        self.setup_window()
    
    def setup_window(self):
        """Configura la ventana de análisis de tasas"""
        # Crear frame principal
        main_frame = ctk.CTkFrame(self.parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        title = ctk.CTkLabel(
            main_frame, 
            text="📊 Análisis de Tasas de Visado", 
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title.pack(pady=10)
        
        # Subtítulo explicativo
        subtitle = ctk.CTkLabel(
            main_frame,
            text="Análisis de tasas pagadas para cálculo de honorarios de ingenieros visadores",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        subtitle.pack(pady=(0, 20))

        # BOTÓN DE SALIDA - CENTRADO Y CON COLORES DEL PROGRAMA
        exit_button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        exit_button_frame.pack(fill="x", padx=10, pady=5)

        btn_exit_top = ctk.CTkButton(
            exit_button_frame,
            text="🏠 Volver al Menú Principal",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.return_callback,
            height=30,
            width=200
        )
        btn_exit_top.pack(anchor="center", pady=5)
        
        # Frame de configuración
        config_frame = ctk.CTkFrame(main_frame)
        config_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Configuración del período
        period_frame = ctk.CTkFrame(config_frame)
        period_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ctk.CTkLabel(period_frame, text="Seleccionar Período de Análisis:", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=5)
        
        # Frame para fechas personalizadas
        date_frame = ctk.CTkFrame(period_frame)
        date_frame.pack(pady=10)
        
        # Fecha de inicio
        ctk.CTkLabel(date_frame, text="Fecha de Inicio:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        # Frame para fecha de inicio
        inicio_frame = ctk.CTkFrame(date_frame)
        inicio_frame.grid(row=0, column=1, padx=5, pady=5)
        
        # Día inicio
        ctk.CTkLabel(inicio_frame, text="Día:").grid(row=0, column=0, padx=2, pady=2)
        self.dia_inicio_var = ctk.StringVar(value="01")
        dias = [f"{i:02d}" for i in range(1, 32)]
        self.dia_inicio_selector = ctk.CTkOptionMenu(inicio_frame, values=dias, variable=self.dia_inicio_var, width=60)
        self.dia_inicio_selector.grid(row=0, column=1, padx=2, pady=2)
        
        # Mes inicio
        ctk.CTkLabel(inicio_frame, text="Mes:").grid(row=0, column=2, padx=2, pady=2)
        self.mes_inicio_var = ctk.StringVar(value=f"{datetime.now().month:02d}")
        meses = [f"{i:02d}" for i in range(1, 13)]
        self.mes_inicio_selector = ctk.CTkOptionMenu(inicio_frame, values=meses, variable=self.mes_inicio_var, width=60)
        self.mes_inicio_selector.grid(row=0, column=3, padx=2, pady=2)
        
        # Año inicio
        ctk.CTkLabel(inicio_frame, text="Año:").grid(row=0, column=4, padx=2, pady=2)
        self.año_inicio_var = ctk.StringVar(value=str(datetime.now().year))
        años = [str(year) for year in range(2020, 2030)]
        self.año_inicio_selector = ctk.CTkOptionMenu(inicio_frame, values=años, variable=self.año_inicio_var, width=80)
        self.año_inicio_selector.grid(row=0, column=5, padx=2, pady=2)
        
        # Fecha de fin
        ctk.CTkLabel(date_frame, text="Fecha de Fin:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        
        # Frame para fecha de fin
        fin_frame = ctk.CTkFrame(date_frame)
        fin_frame.grid(row=1, column=1, padx=5, pady=5)
        
        # Día fin
        ctk.CTkLabel(fin_frame, text="Día:").grid(row=0, column=0, padx=2, pady=2)
        self.dia_fin_var = ctk.StringVar(value=f"{calendar.monthrange(datetime.now().year, datetime.now().month)[1]:02d}")
        self.dia_fin_selector = ctk.CTkOptionMenu(fin_frame, values=dias, variable=self.dia_fin_var, width=60)
        self.dia_fin_selector.grid(row=0, column=1, padx=2, pady=2)
        
        # Mes fin
        ctk.CTkLabel(fin_frame, text="Mes:").grid(row=0, column=2, padx=2, pady=2)
        self.mes_fin_var = ctk.StringVar(value=f"{datetime.now().month:02d}")
        self.mes_fin_selector = ctk.CTkOptionMenu(fin_frame, values=meses, variable=self.mes_fin_var, width=60)
        self.mes_fin_selector.grid(row=0, column=3, padx=2, pady=2)
        
        # Año fin
        ctk.CTkLabel(fin_frame, text="Año:").grid(row=0, column=4, padx=2, pady=2)
        self.año_fin_var = ctk.StringVar(value=str(datetime.now().year))
        self.año_fin_selector = ctk.CTkOptionMenu(fin_frame, values=años, variable=self.año_fin_var, width=80)
        self.año_fin_selector.grid(row=0, column=5, padx=2, pady=2)
        
        # Botones de acceso rápido
        quick_buttons_frame = ctk.CTkFrame(period_frame)
        quick_buttons_frame.pack(pady=10)
        
        ctk.CTkLabel(quick_buttons_frame, text="Accesos rápidos:", font=ctk.CTkFont(size=12, weight="bold")).pack(pady=5)
        
        quick_frame = ctk.CTkFrame(quick_buttons_frame)
        quick_frame.pack(pady=5)
        
        # Botón para mes actual
        btn_mes_actual = ctk.CTkButton(
            quick_frame,
            text="Mes Actual",
            command=self.set_mes_actual,
            width=100,
            height=30
        )
        btn_mes_actual.pack(side=tk.LEFT, padx=5)
        
        # Botón para mes anterior
        btn_mes_anterior = ctk.CTkButton(
            quick_frame,
            text="Mes Anterior",
            command=self.set_mes_anterior,
            width=100,
            height=30
        )
        btn_mes_anterior.pack(side=tk.LEFT, padx=5)
        
        # Botón para últimos 30 días
        btn_30_dias = ctk.CTkButton(
            quick_frame,
            text="Últimos 30 días",
            command=self.set_ultimos_30_dias,
            width=120,
            height=30
        )
        btn_30_dias.pack(side=tk.LEFT, padx=5)
        
        # Botón para año actual
        btn_año_actual = ctk.CTkButton(
            quick_frame,
            text="Año Actual",
            command=self.set_año_actual,
            width=100,
            height=30
        )
        btn_año_actual.pack(side=tk.LEFT, padx=5)
        
        # Botones de acción
        buttons_frame = ctk.CTkFrame(config_frame)
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Botón para generar análisis
        btn_analizar = ctk.CTkButton(
            buttons_frame,
            text="🔍 Generar Análisis",
            font=ctk.CTkFont(size=14),
            command=self.generar_analisis
        )
        btn_analizar.pack(side=tk.LEFT, padx=10, pady=10)
        
        # Botón para exportar a Excel
        self.btn_exportar = ctk.CTkButton(
            buttons_frame,
            text="📤 Exportar a Excel",
            font=ctk.CTkFont(size=14),
            command=self.exportar_excel,
            state="disabled"
        )
        self.btn_exportar.pack(side=tk.LEFT, padx=10, pady=10)
        
        # Botón para cerrar período
        self.btn_cerrar = ctk.CTkButton(
            buttons_frame,
            text="✅ Cerrar Período",
            font=ctk.CTkFont(size=14),
            command=self.cerrar_periodo,
            state="disabled"
        )
        self.btn_cerrar.pack(side=tk.LEFT, padx=10, pady=10)
        
        # Frame de resultados
        self.results_frame = ctk.CTkScrollableFrame(main_frame)
        self.results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Mensaje inicial
        initial_label = ctk.CTkLabel(
            self.results_frame,
            text="Seleccione un período y haga clic en 'Generar Análisis' para comenzar",
            font=ctk.CTkFont(size=14),
            text_color="gray"
        )
        initial_label.pack(pady=50)
        
    
    def set_mes_actual(self):
        """Configura las fechas para el mes actual"""
        hoy = datetime.now()
        primer_dia = 1
        ultimo_dia = calendar.monthrange(hoy.year, hoy.month)[1]
        
        self.dia_inicio_var.set(f"{primer_dia:02d}")
        self.mes_inicio_var.set(f"{hoy.month:02d}")
        self.año_inicio_var.set(str(hoy.year))
        
        self.dia_fin_var.set(f"{ultimo_dia:02d}")
        self.mes_fin_var.set(f"{hoy.month:02d}")
        self.año_fin_var.set(str(hoy.year))
    
    def set_mes_anterior(self):
        """Configura las fechas para el mes anterior"""
        hoy = datetime.now()
        if hoy.month == 1:
            mes_anterior = 12
            año_anterior = hoy.year - 1
        else:
            mes_anterior = hoy.month - 1
            año_anterior = hoy.year
        
        primer_dia = 1
        ultimo_dia = calendar.monthrange(año_anterior, mes_anterior)[1]
        
        self.dia_inicio_var.set(f"{primer_dia:02d}")
        self.mes_inicio_var.set(f"{mes_anterior:02d}")
        self.año_inicio_var.set(str(año_anterior))
        
        self.dia_fin_var.set(f"{ultimo_dia:02d}")
        self.mes_fin_var.set(f"{mes_anterior:02d}")
        self.año_fin_var.set(str(año_anterior))
    
    def set_ultimos_30_dias(self):
        """Configura las fechas para los últimos 30 días"""
        hoy = datetime.now()
        hace_30_dias = hoy - timedelta(days=30)
        
        self.dia_inicio_var.set(f"{hace_30_dias.day:02d}")
        self.mes_inicio_var.set(f"{hace_30_dias.month:02d}")
        self.año_inicio_var.set(str(hace_30_dias.year))
        
        self.dia_fin_var.set(f"{hoy.day:02d}")
        self.mes_fin_var.set(f"{hoy.month:02d}")
        self.año_fin_var.set(str(hoy.year))
    
    def set_año_actual(self):
        """Configura las fechas para todo el año actual"""
        hoy = datetime.now()
        
        self.dia_inicio_var.set("01")
        self.mes_inicio_var.set("01")
        self.año_inicio_var.set(str(hoy.year))
        
        self.dia_fin_var.set("31")
        self.mes_fin_var.set("12")
        self.año_fin_var.set(str(hoy.year))
    
    def generar_analisis(self):
        """Genera el análisis para el período seleccionado"""
        try:
            # Obtener fechas seleccionadas
            try:
                dia_inicio = int(self.dia_inicio_var.get())
                mes_inicio = int(self.mes_inicio_var.get())
                año_inicio = int(self.año_inicio_var.get())
                
                dia_fin = int(self.dia_fin_var.get())
                mes_fin = int(self.mes_fin_var.get())
                año_fin = int(self.año_fin_var.get())
                
                fecha_inicio = datetime(año_inicio, mes_inicio, dia_inicio)
                fecha_fin = datetime(año_fin, mes_fin, dia_fin)
                
                # Validar que la fecha de inicio sea anterior o igual a la fecha de fin
                if fecha_inicio > fecha_fin:
                    messagebox.showerror("Error", "La fecha de inicio debe ser anterior o igual a la fecha de fin")
                    return
                
            except ValueError as e:
                messagebox.showerror("Error", f"Fechas inválidas: {str(e)}")
                return
            
            # Mostrar mensaje de carga
            self.mostrar_mensaje_carga("Generando análisis...")
            
            # Generar análisis con fechas específicas
            self.current_analysis = self.tasas_analyzer.generar_analisis_fechas(
                fecha_inicio, fecha_fin, marcar_como_analizadas=False
            )
            
            if "error" in self.current_analysis:
                messagebox.showerror("Error", self.current_analysis["error"])
                self.limpiar_resultados()
                return
            
            # Mostrar resultados
            self.mostrar_resultados()
            
            # Habilitar botones
            self.btn_exportar.configure(state="normal")
            self.btn_cerrar.configure(state="normal")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar análisis: {str(e)}")
            self.limpiar_resultados()
    
    def mostrar_mensaje_carga(self, mensaje):
        """Muestra un mensaje de carga"""
        # Limpiar frame de resultados
        for widget in self.results_frame.winfo_children():
            widget.destroy()
        
        loading_label = ctk.CTkLabel(
            self.results_frame,
            text=mensaje,
            font=ctk.CTkFont(size=14),
            text_color="blue"
        )
        loading_label.pack(pady=50)
        
        # Actualizar GUI
        self.parent.update()
    
    def mostrar_resultados(self):
        """Muestra los resultados del análisis"""
        # Limpiar frame de resultados
        for widget in self.results_frame.winfo_children():
            widget.destroy()
        
        if not self.current_analysis:
            return
        
        # Título del análisis
        title_frame = ctk.CTkFrame(self.results_frame)
        title_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ctk.CTkLabel(
            title_frame,
            text=f"📊 ANÁLISIS - {self.current_analysis['periodo'].upper()}",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(pady=10)
        
        # Información general
        info_frame = ctk.CTkFrame(self.results_frame)
        info_frame.pack(fill=tk.X, padx=10, pady=10)
        
        info_text = f"""
📈 RESUMEN GENERAL:
• Total de obras con visados: {self.current_analysis['total_obras']}
• Obras pagadas en el período: {self.current_analysis['obras_pagadas']}
• Obras no pagadas (todas): {self.current_analysis['obras_no_pagadas']}
        """
        
        ctk.CTkLabel(info_frame, text=info_text, justify="left", font=ctk.CTkFont(size=12)).pack(pady=10, padx=20)
        
        # Totales por tipo de visado
        self.mostrar_totales_por_tipo()
        
        # Cálculos por ingeniero
        self.mostrar_calculos_por_ingeniero()
        
        # Detalle de obras
        self.mostrar_detalle_obras()
    
    def mostrar_totales_por_tipo(self):
        """Muestra los totales por tipo de visado"""
        totales_frame = ctk.CTkFrame(self.results_frame)
        totales_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ctk.CTkLabel(
            totales_frame,
            text="💰 TOTALES POR TIPO DE VISADO (Solo obras pagadas en el período)",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=10)
        
        # Crear tabla de totales
        table_frame = ctk.CTkFrame(totales_frame)
        table_frame.pack(fill=tk.X, padx=10, pady=10)
        
        headers = ["Tipo", "Total Pagado", "Ingeniero Responsable"]
        for col, header in enumerate(headers):
            ctk.CTkLabel(table_frame, text=header, font=ctk.CTkFont(weight="bold")).grid(
                row=0, column=col, padx=10, pady=5, sticky="w"
            )
        
        totales = self.current_analysis['totales_pagadas']
        tipos_info = [
            ("Gas", totales["gas"]["pagado"], "IMLAUER FERNANDO"),
            ("Salubridad", totales["salubridad"]["pagado"], "IMLAUER FERNANDO"),
            ("Eléctrica", totales["electrica"]["pagado"], "ONETTO JOSE"),
            ("Electromecanica", totales["electromecanica"]["pagado"], "ONETTO JOSE")
        ]
        
        for row, (tipo, total, ingeniero) in enumerate(tipos_info, 1):
            ctk.CTkLabel(table_frame, text=tipo).grid(row=row, column=0, padx=10, pady=3, sticky="w")
            ctk.CTkLabel(table_frame, text=f"${total:,.0f}").grid(row=row, column=1, padx=10, pady=3, sticky="w")
            ctk.CTkLabel(table_frame, text=ingeniero).grid(row=row, column=2, padx=10, pady=3, sticky="w")
    
    def mostrar_calculos_por_ingeniero(self):
        """Muestra los cálculos por ingeniero"""
        ingenieros_frame = ctk.CTkFrame(self.results_frame)
        ingenieros_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ctk.CTkLabel(
            ingenieros_frame,
            text="👷 CÁLCULO DE HONORARIOS POR INGENIERO",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=10)
        
        por_ingeniero = self.current_analysis['por_ingeniero']
        
        for ingeniero, datos in por_ingeniero.items():
            # Frame para cada ingeniero
            ing_frame = ctk.CTkFrame(ingenieros_frame)
            ing_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # Nombre del ingeniero
            ctk.CTkLabel(
                ing_frame,
                text=f"🔧 {ingeniero}",
                font=ctk.CTkFont(size=14, weight="bold")
            ).pack(pady=5)
            
            # Detalles
            detalles_text = f"""
• Total de tasas: ${datos['total']:,.0f}
• Para el Consejo (30%): ${datos['consejo']:,.0f}
• Para el Ingeniero (70%): ${datos['ingeniero']:,.0f}
• Tipos de visado: {', '.join(datos['tipos']).title()}
            """
            
            ctk.CTkLabel(ing_frame, text=detalles_text, justify="left", font=ctk.CTkFont(size=11)).pack(
                pady=5, padx=20
            )
        
        # Total general
        total_frame = ctk.CTkFrame(ingenieros_frame)
        total_frame.pack(fill=tk.X, padx=10, pady=10)
        
        total_general = sum(d['total'] for d in por_ingeniero.values())
        total_consejo = sum(d['consejo'] for d in por_ingeniero.values())
        total_ingenieros = sum(d['ingeniero'] for d in por_ingeniero.values())
        
        total_text = f"""
🏛️ TOTALES GENERALES:
• Total de todas las tasas: ${total_general:,.0f}
• Total para el Consejo (30%): ${total_consejo:,.0f}
• Total para Ingenieros (70%): ${total_ingenieros:,.0f}
        """
        
        ctk.CTkLabel(
            total_frame,
            text=total_text,
            justify="left",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(pady=10, padx=20)
    
    def mostrar_detalle_obras(self):
        """Muestra el detalle de las obras analizadas"""
        detalle_frame = ctk.CTkFrame(self.results_frame)
        detalle_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ctk.CTkLabel(
            detalle_frame,
            text="📋 DETALLE DE OBRAS ANALIZADAS",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=10)
        
        # Crear pestañas para obras pagadas y no pagadas
        tab_view = ctk.CTkTabview(detalle_frame)
        tab_view.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Pestaña de obras pagadas
        tab_pagadas = tab_view.add(f"Pagadas en período ({self.current_analysis['obras_pagadas']})")
        self.crear_tabla_obras(tab_pagadas, self.current_analysis['obras_pagadas_detalle'], "pagadas")
        
        # Pestaña de obras no pagadas
        tab_no_pagadas = tab_view.add(f"No Pagadas - Todas ({self.current_analysis['obras_no_pagadas']})")
        self.crear_tabla_obras(tab_no_pagadas, self.current_analysis['obras_no_pagadas_detalle'], "no_pagadas")
    
    def crear_tabla_obras(self, parent, obras, tipo):
        """Crea una tabla con el detalle de las obras"""
        if not obras:
            mensaje = f"No hay obras {tipo} " + ("en este período" if tipo == "pagadas" else "pendientes")
            ctk.CTkLabel(parent, text=mensaje, font=ctk.CTkFont(size=12)).pack(pady=20)
            return
        
        # Frame scrollable para la tabla
        table_frame = ctk.CTkScrollableFrame(parent)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Encabezados
        if tipo == "pagadas":
            headers = ["Fecha Entrada", "Fecha Salida", "Profesional", "Comitente", "Gas", "Salubridad", "Eléctrica", "Electromecanica", "Total"]
        else:
            headers = ["Fecha Entrada", "Profesional", "Comitente", "Gas", "Salubridad", "Eléctrica", "Electromecanica", "Total"]
        
        for col, header in enumerate(headers):
            ctk.CTkLabel(table_frame, text=header, font=ctk.CTkFont(weight="bold")).grid(
                row=0, column=col, padx=5, pady=5, sticky="w"
            )
        
        # Datos
        for row, obra in enumerate(obras, 1):
            gas = float(obra.get("visado_gas", 0) or 0)
            salubridad = float(obra.get("visado_salubridad", 0) or 0)
            electrica = float(obra.get("visado_electrica", 0) or 0)
            electromecanica = float(obra.get("visado_electromecanica", 0) or 0)
            total = gas + salubridad + electrica + electromecanica
            
            if tipo == "pagadas":
                datos = [
                    obra.get("fecha", ""),
                    obra.get("fecha_salida", ""),
                    obra.get("nombre_profesional", "")[:20] + "..." if len(obra.get("nombre_profesional", "")) > 20 else obra.get("nombre_profesional", ""),
                    obra.get("nombre_comitente", "")[:20] + "..." if len(obra.get("nombre_comitente", "")) > 20 else obra.get("nombre_comitente", ""),
                    f"${gas:,.0f}" if gas > 0 else "-",
                    f"${salubridad:,.0f}" if salubridad > 0 else "-",
                    f"${electrica:,.0f}" if electrica > 0 else "-",
                    f"${electromecanica:,.0f}" if electromecanica > 0 else "-",
                    f"${total:,.0f}"
                ]
            else:
                datos = [
                    obra.get("fecha", ""),
                    obra.get("nombre_profesional", "")[:20] + "..." if len(obra.get("nombre_profesional", "")) > 20 else obra.get("nombre_profesional", ""),
                    obra.get("nombre_comitente", "")[:20] + "..." if len(obra.get("nombre_comitente", "")) > 20 else obra.get("nombre_comitente", ""),
                    f"${gas:,.0f}" if gas > 0 else "-",
                    f"${salubridad:,.0f}" if salubridad > 0 else "-",
                    f"${electrica:,.0f}" if electrica > 0 else "-",
                    f"${electromecanica:,.0f}" if electromecanica > 0 else "-",
                    f"${total:,.0f}"
                ]
            
            for col, dato in enumerate(datos):
                ctk.CTkLabel(table_frame, text=dato, font=ctk.CTkFont(size=10)).grid(
                    row=row, column=col, padx=5, pady=2, sticky="w"
                )
    
    def exportar_excel(self):
        """Exporta el análisis a Excel"""
        if not self.current_analysis:
            messagebox.showwarning("Advertencia", "No hay análisis para exportar")
            return
        
        try:
            # Generar nombre de archivo sugerido
            periodo = self.current_analysis['periodo'].replace(' ', '_').replace('/', '-').replace('á', 'a').replace('é', 'e')
            nombre_sugerido = f"Analisis_Tasas_{periodo}.xlsx"
            
            # Abrir diálogo para guardar archivo con manejo de errores
            try:
                archivo = filedialog.asksaveasfilename(
                    title="Guardar análisis de tasas",
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                    initialfile=nombre_sugerido
                )
            except Exception as dialog_error:
                print(f"Error en el diálogo: {dialog_error}")
                # Si falla el diálogo con nombre sugerido, intentar sin él
                archivo = filedialog.asksaveasfilename(
                    title="Guardar análisis de tasas",
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
                )
            
            if archivo:
                # Exportar
                resultado = self.tasas_analyzer.exportar_a_excel(self.current_analysis, archivo)
                
                if resultado:
                    messagebox.showinfo("Éxito", f"Análisis exportado correctamente a:\n{archivo}")
                else:
                    messagebox.showerror("Error", "No se pudo exportar el análisis")
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar: {str(e)}")
    
    def cerrar_periodo(self):
        """Cierra el período marcando las obras como analizadas"""
        if not self.current_analysis:
            messagebox.showwarning("Advertencia", "No hay análisis para cerrar")
            return
        
        # Confirmar acción
        respuesta = messagebox.askyesno(
            "Confirmar cierre de período",
            f"¿Está seguro de que desea cerrar el período {self.current_analysis['periodo']}?\n\n"
            f"Esto marcará {self.current_analysis['obras_pagadas']} obras como 'ya analizadas' "
            f"y no aparecerán en futuros análisis.\n\n"
            "Esta acción NO se puede deshacer."
        )
        
        if respuesta:
            try:
                # Regenerar análisis marcando como analizadas
                fecha_inicio = self.current_analysis['fecha_inicio']
                fecha_fin = self.current_analysis['fecha_fin']
                
                self.current_analysis = self.tasas_analyzer.generar_analisis_fechas(
                    fecha_inicio, fecha_fin, marcar_como_analizadas=True
                )
                
                messagebox.showinfo(
                    "Período cerrado",
                    f"El período {self.current_analysis['periodo']} ha sido cerrado exitosamente.\n"
                    f"Las obras pagadas han sido marcadas como analizadas."
                )
                
                # Deshabilitar botón de cerrar
                self.btn_cerrar.configure(state="disabled")
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al cerrar período: {str(e)}")
    
    def limpiar_resultados(self):
        """Limpia los resultados y deshabilita botones"""
        # Limpiar frame de resultados
        for widget in self.results_frame.winfo_children():
            widget.destroy()
        
        # Mensaje inicial
        initial_label = ctk.CTkLabel(
            self.results_frame,
            text="Seleccione un período y haga clic en 'Generar Análisis' para comenzar",
            font=ctk.CTkFont(size=14),
            text_color="gray"
        )
        initial_label.pack(pady=50)
        
        # Deshabilitar botones
        self.btn_exportar.configure(state="disabled")
        self.btn_cerrar.configure(state="disabled")
        
        # Limpiar análisis actual
        self.current_analysis = None