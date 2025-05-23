#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Sistema de Gestión de Trabajos Profesionales - CPIM
-------------------------------------------------
Aplicación para registrar y organizar trabajos profesionales 
que ingresan al Consejo Profesional de Ingeniería de Misiones.

Fecha: Abril 2025
"""

import os
import sys
import customtkinter as ctk
from modules.gui import App
from config import ensure_directories


def main():
    """Función principal del programa"""
    # Configurar tema de la aplicación
    ctk.set_appearance_mode("System")
    # Modos: "System" (default), "Dark", "Light"
    ctk.set_default_color_theme("blue")  # Temas: "blue" (default), "dark-blue", "green"
    
    # Asegurar que las carpetas necesarias existen
    ensure_directories()
    
    # Iniciar la aplicación
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()