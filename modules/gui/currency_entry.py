import tkinter as tk
import customtkinter as ctk
import re


class CurrencyEntry(ctk.CTkFrame):
    """Campo de entrada para moneda con formato argentino"""
    def __init__(self, parent, textvariable=None, width=200, height=30, **kwargs):
        super().__init__(parent, width=width, height=height)
        
        self.parent = parent
        self.raw_value = 0.0  # Valor numérico sin formato
        
        # Variable interna para el texto formateado
        self.formatted_var = ctk.StringVar()
        
        # Variable externa (si se proporciona)
        self.external_var = textvariable
        
        # Crear el widget de entrada
        self.entry = ctk.CTkEntry(self, textvariable=self.formatted_var, width=width, height=height, **kwargs)
        self.entry.pack(fill="x", expand=True)
        
        # Configurar eventos
        self.formatted_var.trace_add("write", self.on_text_changed)
        self.entry.bind("<FocusOut>", self.on_focus_out)
        self.entry.bind("<KeyPress>", self.on_key_press)
        
        # Configurar placeholder
        self.entry.configure(placeholder_text="$0,00")
    
    def on_key_press(self, event):
        """Maneja las teclas presionadas"""
        # Permitir teclas de control
        control_keys = ['BackSpace', 'Delete', 'Left', 'Right', 'Home', 'End', 'Tab']
        if event.keysym in control_keys:
            return
        
        # Permitir solo números, coma y punto
        if not re.match(r'[0-9.,]', event.char):
            return "break"
    
    def on_text_changed(self, *args):
        """Se ejecuta cuando cambia el texto"""
        try:
            current_text = self.formatted_var.get()
            
            # Si está vacío, limpiar todo
            if not current_text:
                self.raw_value = 0.0
                if self.external_var:
                    self.external_var.set("")
                return
            
            # Extraer solo números del texto
            numbers_only = re.sub(r'[^\d,]', '', current_text)
            
            # Si no hay números, mantener como 0
            if not numbers_only:
                self.raw_value = 0.0
                if self.external_var:
                    self.external_var.set("")
                return
            
            # LÓGICA CORREGIDA: Manejar números enteros y decimales correctamente
            if ',' in numbers_only:
                # Si hay coma, separar enteros y decimales
                parts = numbers_only.split(',')
                integer_part = parts[0] if parts[0] else "0"
                decimal_part = parts[1][:2] if len(parts) > 1 else "00"
                
                # Formar el número con decimales
                self.raw_value = float(f"{integer_part}.{decimal_part}")
            else:
                # CORRECCIÓN: Si no hay coma, es un número entero
                # No agregar decimales automáticamente
                self.raw_value = float(numbers_only)
            
            # Actualizar variable externa con el valor numérico
            if self.external_var:
                self.external_var.set(str(self.raw_value))
                    
        except Exception as e:
            print(f"Error en formato de moneda: {e}")
    
    def on_focus_out(self, event):
        """Cuando pierde el foco, formatear completamente"""
        self.format_display()
    
    def format_display(self):
        """Formatea el valor para mostrar"""
        try:
            if self.raw_value == 0:
                self.formatted_var.set("")
                return
            
            # CORRECCIÓN: Verificar si el número es entero o tiene decimales
            if self.raw_value == int(self.raw_value):
                # Es un número entero, mostrar sin decimales
                formatted = f"${int(self.raw_value):,}"
            else:
                # Tiene decimales, mostrar con 2 decimales
                formatted = f"${self.raw_value:,.2f}"
            
            # Cambiar punto por coma para decimales y coma por punto para miles
            formatted = formatted.replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.')
            
            # Temporalmente desconectar el trace para evitar bucle
            trace_info = self.formatted_var.trace_info()
            if trace_info:
                self.formatted_var.trace_remove("write", trace_info[0][1])
            self.formatted_var.set(formatted)
            self.formatted_var.trace_add("write", self.on_text_changed)
            
        except Exception as e:
            print(f"Error al formatear: {e}")
    
    def set(self, value):
        """Establece el valor del campo"""
        try:
            if isinstance(value, str):
                if not value or value.strip() == "":
                    self.raw_value = 0.0
                else:
                    # CORRECCIÓN: Limpiar correctamente el string
                    # Remover $ y espacios, pero mantener puntos y comas
                    clean_value = value.replace("$", "").replace(" ", "")
                    
                    # Si tiene formato argentino (punto para miles, coma para decimales)
                    if "." in clean_value and "," in clean_value:
                        # Formato: 12.000,50 -> convertir a 12000.50
                        parts = clean_value.split(",")
                        integer_part = parts[0].replace(".", "")  # Remover puntos de miles
                        decimal_part = parts[1] if len(parts) > 1 else "00"
                        clean_value = f"{integer_part}.{decimal_part}"
                    elif "," in clean_value and "." not in clean_value:
                        # Solo coma decimal: 12000,50 -> 12000.50
                        clean_value = clean_value.replace(",", ".")
                    elif "." in clean_value and "," not in clean_value:
                        # Solo punto: podría ser miles o decimal
                        # Si hay más de 3 dígitos después del punto, es separador de miles
                        parts = clean_value.split(".")
                        if len(parts) == 2 and len(parts[1]) > 2:
                            # Es separador de miles: 12.000 -> 12000
                            clean_value = clean_value.replace(".", "")
                        # Si tiene 1-2 dígitos después del punto, es decimal
                    
                    self.raw_value = float(clean_value) if clean_value else 0.0
            else:
                self.raw_value = float(value) if value else 0.0
            
            self.format_display()
            
        except (ValueError, TypeError) as e:
            print(f"Error al establecer valor: {e}")
            self.raw_value = 0.0
            self.formatted_var.set("")
    
    def get(self):
        """Retorna el valor numérico como string"""
        # CORRECCIÓN: Devolver el valor raw sin formatear
        if self.raw_value == 0:
            return ""
        
        # Si es un número entero, devolverlo sin decimales
        if self.raw_value == int(self.raw_value):
            return str(int(self.raw_value))
        else:
            return str(self.raw_value)
    
    def get_float(self):
        """Retorna el valor numérico como float"""
        return self.raw_value
    
    def insert(self, index, text):
        """Compatibilidad con Entry - inserta texto"""
        self.set(text)
    
    def configure(self, **kwargs):
        """Pasa la configuración al entry interno"""
        self.entry.configure(**kwargs)
    
    def grid(self, **kwargs):
        """Override grid para el frame"""
        super().grid(**kwargs)
    
    def pack(self, **kwargs):
        """Override pack para el frame"""
        super().pack(**kwargs)

