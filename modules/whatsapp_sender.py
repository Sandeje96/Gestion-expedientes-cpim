import webbrowser
import urllib.parse
import subprocess
import os
import time
import tempfile
from config import WHATSAPP_WEB_URL


class WhatsAppSender:
    def __init__(self):
        pass
    
    def format_phone_number(self, phone):
        """Formatea el número de teléfono para WhatsApp"""
        if not phone:
            return None
        
        # Remover espacios, guiones y paréntesis
        clean_phone = ''.join(filter(str.isdigit, phone))
        
        # Si no tiene código de país, agregar Argentina (+54)
        if len(clean_phone) == 10:  # Número argentino sin código de país
            clean_phone = "54" + clean_phone
        elif len(clean_phone) == 11 and clean_phone.startswith("0"):
            # Número argentino con 0 inicial, remover el 0 y agregar 54
            clean_phone = "54" + clean_phone[1:]
        
        return clean_phone
    
    def send_whatsapp_message(self, phone, message):
        """Muestra el mensaje en una ventana para copiarlo manualmente a WhatsApp Web"""
        try:
            formatted_phone = self.format_phone_number(phone)
            if not formatted_phone:
                return False
            
            # Importar tkinter para mostrar ventana
            import tkinter as tk
            from tkinter import messagebox, Toplevel, Text, Frame, Button
            import customtkinter as ctk
            
            # Crear ventana emergente
            popup = Toplevel()
            popup.title(f"Mensaje para {phone}")
            popup.geometry("550x450")  # Aumentar tamaño mínimo
            popup.minsize(450, 350)    # Establecer tamaño mínimo
            popup.resizable(True, True)
            
            # Hacer la ventana modal y siempre al frente
            popup.transient()
            popup.grab_set()
            popup.attributes('-topmost', True)
            
            # Frame principal con padding
            main_frame = Frame(popup)
            main_frame.pack(fill="both", expand=True, padx=15, pady=15)
            
            # Título
            title_label = tk.Label(main_frame, text=f"📱 Mensaje para: {phone}", 
                                 font=("Arial", 12, "bold"))
            title_label.pack(pady=(0, 10))
            
            # Área de texto con el mensaje - usando grid para mejor control
            text_frame = Frame(main_frame)
            text_frame.pack(fill="both", expand=True, pady=(0, 15))
            
            text_area = Text(text_frame, wrap="word", font=("Arial", 10), height=10)
            
            # Agregar scrollbar
            scrollbar = tk.Scrollbar(text_frame, orient="vertical", command=text_area.yview)
            text_area.configure(yscrollcommand=scrollbar.set)
            
            text_area.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            text_area.insert("1.0", message)
            text_area.config(state="normal")  # Permitir selección
            
            # Seleccionar todo el texto automáticamente
            text_area.tag_add("sel", "1.0", "end-1c")
            text_area.mark_set("insert", "1.0")
            text_area.focus_set()
            
            # Frame para botones - SIEMPRE VISIBLE al final
            button_frame = Frame(main_frame)
            button_frame.pack(fill="x", side="bottom", pady=(10, 0))
            
            # Función para copiar al portapapeles
            def copy_to_clipboard():
                popup.clipboard_clear()
                popup.clipboard_append(message)
                popup.update()  # Mantener en portapapeles
                messagebox.showinfo("Copiado", "Mensaje copiado al portapapeles.\nPégalo en WhatsApp Web.")
            
            # Función para abrir WhatsApp Web
            def open_whatsapp():
                encoded_message = urllib.parse.quote(message)
                url = f"{WHATSAPP_WEB_URL}?phone={formatted_phone}&text={encoded_message}"
                webbrowser.open(url)
                popup.destroy()
            
            # Botones con mejor distribución
            btn_copy = Button(button_frame, text="📋 Copiar al Portapapeles", 
                             command=copy_to_clipboard, bg="#4CAF50", fg="white", 
                             font=("Arial", 9, "bold"), height=2, width=20)
            btn_copy.pack(side="left", padx=(0, 5), fill="x", expand=True)
            
            btn_open = Button(button_frame, text="🌐 Abrir WhatsApp Web", 
                             command=open_whatsapp, bg="#25D366", fg="white", 
                             font=("Arial", 9, "bold"), height=2, width=20)
            btn_open.pack(side="left", padx=5, fill="x", expand=True)
            
            btn_close = Button(button_frame, text="❌ Cerrar", 
                              command=popup.destroy, bg="#f44336", fg="white", 
                              font=("Arial", 9, "bold"), height=2, width=15)
            btn_close.pack(side="right", padx=(5, 0))
            
            # Instrucciones - SIEMPRE VISIBLE
            instructions_frame = Frame(main_frame)
            instructions_frame.pack(fill="x", side="bottom", pady=(5, 0))
            
            instructions = tk.Label(instructions_frame, 
                                  text="💡 Copia el mensaje y pégalo en tu WhatsApp Web abierto", 
                                  font=("Arial", 9), fg="gray")
            instructions.pack()
            
            # Centrar la ventana
            popup.update_idletasks()
            x = (popup.winfo_screenwidth() // 2) - (popup.winfo_width() // 2)
            y = (popup.winfo_screenheight() // 2) - (popup.winfo_height() // 2)
            popup.geometry(f"+{x}+{y}")
            
            # Esperar a que se cierre la ventana
            popup.wait_window()
            
            return True
            
        except Exception as e:
            print(f"Error al mostrar mensaje WhatsApp: {e}")
            return False
    
    def send_whatsapp_message_simple(self, phone, message):
        """Versión simplificada que solo copia al portapapeles"""
        try:
            formatted_phone = self.format_phone_number(phone)
            if not formatted_phone:
                return False
            
            # Crear el mensaje completo para copiar
            full_message = f"Número: {phone}\nMensaje:\n{message}"
            
            # Intentar copiar al portapapeles
            try:
                import pyperclip
                pyperclip.copy(full_message)
                print(f"Mensaje copiado al portapapeles para {phone}")
                return True
            except ImportError:
                # Si no está pyperclip, usar método alternativo
                return self.send_whatsapp_message(phone, message)
                
        except Exception as e:
            print(f"Error al preparar WhatsApp: {e}")
            return False
    
    def create_payment_message(self, nombre_profesional, nombre_comitente, tasa_sellado, total_visados):
        """Crea el mensaje de notificación de pagos"""
        message = f"📋 *CPIM - Notificación de Tasas*\n\n"
        message += f"*Profesional:* {nombre_profesional}\n"
        message += f"*Comitente:* {nombre_comitente}\n\n"
        
        # Mostrar tasa de sellado si existe
        if tasa_sellado and str(tasa_sellado).strip():
            message += f"💰 *Tasa de Sellado:* ${tasa_sellado}\n"
        
        # Mostrar total de visados si existe
        if total_visados > 0:
            message += f"🔧 *Total Visados:* ${total_visados}\n"
        
        # Calcular y mostrar total solo si hay valores
        total_amount = 0
        if tasa_sellado and str(tasa_sellado).strip():
            try:
                total_amount += float(tasa_sellado)
            except (ValueError, TypeError):
                pass
        
        total_amount += total_visados
        
        if total_amount > 0:
            message += f"\n*Total a pagar:* ${total_amount}\n"
        
        message += "\n¡Gracias por su atención!"
        
        return message
    
    def send_payment_notifications(self, obra_data, tasa_sellado, visados, use_simple_method=False):
        """Envía notificaciones de pago a los números de WhatsApp"""
        results = []
        
        # Calcular total de visados
        total_visados = 0
        for visado in visados.values():
            if visado and str(visado).strip():
                try:
                    total_visados += float(visado)
                except (ValueError, TypeError):
                    pass
        
        # Verificar si hay algo que notificar
        has_sellado = tasa_sellado and str(tasa_sellado).strip()
        has_visados = total_visados > 0
        
        if not has_sellado and not has_visados:
            return []  # No hay nada que notificar
        
        # Crear mensaje
        message = self.create_payment_message(
            obra_data.get("nombre_profesional", ""),
            obra_data.get("nombre_comitente", ""),
            tasa_sellado if has_sellado else "",
            total_visados
        )
        
        # Enviar a profesional
        phone_profesional = obra_data.get("whatsapp_profesional")
        if phone_profesional and phone_profesional.strip():
            if use_simple_method:
                success = self.send_whatsapp_message_simple(phone_profesional, message)
            else:
                success = self.send_whatsapp_message(phone_profesional, message)
            results.append(("Profesional", phone_profesional, success))
        
        # Enviar a tramitador
        phone_tramitador = obra_data.get("whatsapp_tramitador")
        if phone_tramitador and phone_tramitador.strip():
            if use_simple_method:
                success = self.send_whatsapp_message_simple(phone_tramitador, message)
            else:
                success = self.send_whatsapp_message(phone_tramitador, message)
            results.append(("Tramitador", phone_tramitador, success))
        
        return results