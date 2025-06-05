import tkinter as tk
import customtkinter as ctk
from tkinter import TclError


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
            # Trigger any trace callbacks after selection
            self.entry_var.set(self.entry_var.get())  # Force trace to execute
    
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