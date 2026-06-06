from pathlib import Path
import win32com.client
import os
import sys
import shutil
import time
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading

class StreamBatGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Stream Bat App")
        self.root.geometry("800x600")
        
        # Inicializar diccionarios
        self.obs_path = None
        self.steam_dictionary = {}
        self.program_shortcut_dictionary = {}
        self.web_dictionary = {}
        self.ms_store_dictionary = {}
        self.obs_shortcut = None
        
        # Establecer estilos
        self.setup_styles()
        
        # Crear contenedor principal
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar pesos de grid
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)
        
        # Crear encabezado
        self.create_header()
        
        # Crear notebook para pestañas
        self.create_notebook()
        
        # Crear barra de estado
        self.create_status_bar()
        
        # Centrar ventana
        self.center_window()

        self.refresh_view()
        
    def setup_styles(self):
        style = ttk.Style()
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'))
        style.configure('Header.TLabel', font=('Arial', 12, 'bold'))
        style.configure('Success.TLabel', foreground='green')
        style.configure('Error.TLabel', foreground='red')
        style.configure('Accent.TButton', font=('Arial', 10, 'bold'))

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def create_header(self):
        header_frame = ttk.Frame(self.main_frame)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        title = ttk.Label(header_frame, text="🦇 Stream Bat App", style='Title.TLabel')
        title.pack(side=tk.LEFT)
        
        subtitle = ttk.Label(header_frame, 
                            text="¡Fabrique un archivo batch (.bat) para iniciar todos sus programas en un sólo clic!")
        subtitle.pack(side=tk.LEFT, padx=(10, 0))

        author_label = ttk.Label(header_frame, text="Por: Timothy Figueroa, con\nayuda de Deepseek-V3", font=('Arial', 8), foreground="gray")
        author_label.pack(side=tk.RIGHT)

    def create_notebook(self):
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Crear pestañas
        self.tab_obs = ttk.Frame(self.notebook)
        self.tab_programs = ttk.Frame(self.notebook)
        self.tab_steam = ttk.Frame(self.notebook)
        self.tab_ms_store = ttk.Frame(self.notebook)
        self.tab_web = ttk.Frame(self.notebook)
        self.tab_view = ttk.Frame(self.notebook)
        self.tab_build = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_obs, text="📷 OBS")
        self.notebook.add(self.tab_programs, text="💻 Programas")
        self.notebook.add(self.tab_steam, text="🎮 Steam")
        self.notebook.add(self.tab_ms_store, text="🏪 Tienda MS")
        self.notebook.add(self.tab_web, text="🌐 Páginas Web")
        self.notebook.add(self.tab_view, text="👁️ Ver Lista")
        self.notebook.add(self.tab_build, text="⚙️ Construir")

        # Poblar pestañas
        self.create_obs_tab()
        self.create_programs_tab()
        self.create_steam_tab()
        self.create_ms_store_tab()
        self.create_web_tab()
        self.create_view_tab()
        self.create_build_tab()

    def create_obs_tab(self):
        # Marco OBS
        obs_frame = ttk.LabelFrame(self.tab_obs, text="Configuración OBS Studio", padding="10")
        obs_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        ttk.Label(obs_frame, text="Agregue su ejecutable de OBS Studio:").pack(anchor=tk.W, pady=(0, 10))
        
        # Etiqueta de instrucción
        ttk.Label(obs_frame, text="Puede buscar el archivo o pegar la ruta manualmente:").pack(anchor=tk.W, pady=(0, 5))
        
        # Marco de entrada de ruta
        path_frame = ttk.Frame(obs_frame)
        path_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.obs_path_var = tk.StringVar()
        self.obs_path_entry = ttk.Entry(path_frame, textvariable=self.obs_path_var, width=50)
        self.obs_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        # Marco de botones
        button_frame = ttk.Frame(path_frame)
        button_frame.pack(side=tk.RIGHT)
        
        ttk.Button(button_frame, text="Buscar...", command=self.browse_obs, width=10).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Establecer Ruta", command=self.set_obs_path, width=15).pack(side=tk.RIGHT)
        
        # Estado actual de OBS
        self.obs_status_frame = ttk.LabelFrame(obs_frame, text="OBS Actual", padding="10")
        self.obs_status_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Crear marco para etiqueta de estado y botón limpiar
        status_content_frame = ttk.Frame(self.obs_status_frame)
        status_content_frame.pack(fill=tk.X, expand=True)
        
        self.obs_status_label = ttk.Label(status_content_frame, text="No hay OBS configurado")
        self.obs_status_label.pack(side=tk.LEFT, anchor=tk.W, fill=tk.X, expand=True)
        
        # Agregar botón Limpiar OBS
        self.clear_obs_button = ttk.Button(status_content_frame, text="Eliminar Ruta de OBS", 
                                           command=self.clear_obs_path, width=20)
        self.clear_obs_button.pack(side=tk.RIGHT)
        
        # Inicialmente ocultar botón limpiar (solo mostrar cuando OBS está configurado)
        self.clear_obs_button.pack_forget()
        
        # Actualizar estado
        self.update_obs_status()

    def clear_obs_path(self):
        """Borrar la ruta de OBS y restablecer a no configurado."""
        if self.obs_path:
            # Pedir confirmación
            response = messagebox.askyesno("Borrar Ruta de OBS", 
                                            "¿Está seguro/a de que quiere borrar la ruta de OBS?\n\n"
                                            "Esto eliminará OBS de su lista de lanzamiento.")
            if response:
                # Limpiar ruta
                self.obs_path = None
                self.obs_path_var.set('')
                
                # Actualizar estado
                self.update_obs_status()
                
                # Auto-refrescar para actualizar la lista de vista
                self.auto_refresh()
                
                self.update_status("Ruta de OBS borrada")
        else:
            messagebox.showinfo("Información", "No hay ruta de OBS configurada actualmente.")

    def create_programs_tab(self):
        # Marco de programas
        programs_frame = ttk.Frame(self.tab_programs)
        programs_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Sección agregar programa
        add_frame = ttk.LabelFrame(programs_frame, text="Agregar Programa", padding="10")
        add_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Entrada de nombre
        ttk.Label(add_frame, text="Nombre del Programa:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.prog_name_var = tk.StringVar()
        ttk.Entry(add_frame, textvariable=self.prog_name_var, width=30).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 5), padx=(5, 0))
        
        # Entrada de ruta
        ttk.Label(add_frame, text="Ruta del Ejecutable:").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        path_frame = ttk.Frame(add_frame)
        path_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 5), padx=(5, 0))
        
        self.prog_path_var = tk.StringVar()
        self.prog_path_entry = ttk.Entry(path_frame, textvariable=self.prog_path_var, width=40)
        self.prog_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        ttk.Button(path_frame, text="Buscar...", command=self.browse_program).pack(side=tk.RIGHT)
        
        # Botón agregar
        ttk.Button(add_frame, text="Agregar Programa", command=self.add_program).grid(row=2, column=0, columnspan=2, pady=(10, 0))
        
        # Configurar pesos de grid
        add_frame.columnconfigure(1, weight=1)
        
        # Lista de programas
        list_frame = ttk.LabelFrame(programs_frame, text="Programas Agregados", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Treeview para programas
        columns = ('name', 'path')
        self.programs_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=8)
        
        self.programs_tree.heading('name', text='Nombre del Programa')
        self.programs_tree.heading('path', text='Ruta del Ejecutable')
        
        self.programs_tree.column('name', width=150)
        self.programs_tree.column('path', width=400)
        
        # Barra de desplazamiento
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.programs_tree.yview)
        self.programs_tree.configure(yscrollcommand=scrollbar.set)
        
        self.programs_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Botón eliminar
        ttk.Button(list_frame, text="Eliminar Seleccionado", command=self.remove_program).pack(pady=(5, 0))

    def create_steam_tab(self):
        # Marco Steam
        steam_frame = ttk.Frame(self.tab_steam)
        steam_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Sección agregar programa Steam
        add_frame = ttk.LabelFrame(steam_frame, text="Agregar Juego/Programa de Steam", padding="10")
        add_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Entrada de nombre
        ttk.Label(add_frame, text="Nombre del Juego/Programa:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.steam_name_var = tk.StringVar()
        ttk.Entry(add_frame, textvariable=self.steam_name_var, width=30).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 5), padx=(5, 0))
        
        # Entrada de URL
        ttk.Label(add_frame, text="URL de Steam:").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        self.steam_url_var = tk.StringVar()
        url_entry = ttk.Entry(add_frame, textvariable=self.steam_url_var, width=40)
        url_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 5), padx=(5, 0))
        
        # Ejemplo
        example = ttk.Label(add_frame, text="Ejemplo: steam://rungameid/123456", foreground="gray")
        example.grid(row=2, column=1, sticky=tk.W, pady=(0, 10), padx=(5, 0))
        
        # Botón agregar
        ttk.Button(add_frame, text="Agregar Juego/Programa de Steam", command=self.add_steam).grid(row=3, column=0, columnspan=2)
        
        # Configurar pesos de grid
        add_frame.columnconfigure(1, weight=1)
        
        # Lista de juegos Steam
        list_frame = ttk.LabelFrame(steam_frame, text="Juegos/Programas de Steam Agregados", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Treeview para juegos Steam
        columns = ('name', 'url')
        self.steam_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=8)
        
        self.steam_tree.heading('name', text='Nombre del Juego/Programa')
        self.steam_tree.heading('url', text='URL de Steam')
        
        self.steam_tree.column('name', width=150)
        self.steam_tree.column('url', width=400)
        
        # Barra de desplazamiento
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.steam_tree.yview)
        self.steam_tree.configure(yscrollcommand=scrollbar.set)
        
        self.steam_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Botón eliminar
        ttk.Button(list_frame, text="Eliminar Seleccionado", command=self.remove_steam).pack(pady=(5, 0))

    def create_ms_store_tab(self):
        ms_frame = ttk.Frame(self.tab_ms_store)
        ms_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Sección de búsqueda
        search_frame = ttk.LabelFrame(ms_frame, text="Buscar una App de la Tienda Microsoft", padding="10")
        search_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(search_frame, text="Escriba el nombre de la app y haga clic en Buscar:").pack(anchor=tk.W, pady=(0, 5))

        input_row = ttk.Frame(search_frame)
        input_row.pack(fill=tk.X)

        self.ms_search_var = tk.StringVar()
        self.ms_search_entry = ttk.Entry(input_row, textvariable=self.ms_search_var, width=35)
        self.ms_search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.ms_search_entry.bind('<Return>', lambda e: self.search_ms_store_apps())

        self.ms_search_button = ttk.Button(input_row, text="Buscar", command=self.search_ms_store_apps, width=10)
        self.ms_search_button.pack(side=tk.LEFT)

        self.ms_search_status = ttk.Label(search_frame, text="", foreground="gray", font=('Arial', 9))
        self.ms_search_status.pack(anchor=tk.W, pady=(5, 0))

        # Sección de resultados
        results_frame = ttk.LabelFrame(ms_frame, text="Resultados — seleccione una app y haga clic en Agregar", padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        res_columns = ('name', 'aumid')
        self.ms_results_tree = ttk.Treeview(results_frame, columns=res_columns, show='headings', height=5)
        self.ms_results_tree.heading('name', text='Nombre de la App')
        self.ms_results_tree.heading('aumid', text='ID de la App (AUMID)')
        self.ms_results_tree.column('name', width=160)
        self.ms_results_tree.column('aumid', width=370)

        res_scroll = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.ms_results_tree.yview)
        self.ms_results_tree.configure(yscrollcommand=res_scroll.set)
        self.ms_results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        res_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Button(results_frame, text="Agregar App Seleccionada", command=self.add_ms_store_from_results).pack(pady=(5, 0))

        # Lista de apps agregadas
        list_frame = ttk.LabelFrame(ms_frame, text="Apps de la Tienda Microsoft Agregadas", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True)

        columns = ('name', 'aumid')
        self.ms_store_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=4)
        self.ms_store_tree.heading('name', text='Nombre de la App')
        self.ms_store_tree.heading('aumid', text='ID de la App (AUMID)')
        self.ms_store_tree.column('name', width=160)
        self.ms_store_tree.column('aumid', width=370)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.ms_store_tree.yview)
        self.ms_store_tree.configure(yscrollcommand=scrollbar.set)
        self.ms_store_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Button(list_frame, text="Eliminar Seleccionado", command=self.remove_ms_store).pack(pady=(5, 0))

    def create_web_tab(self):
        # Marco Web
        web_frame = ttk.Frame(self.tab_web)
        web_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Sección agregar página web
        add_frame = ttk.LabelFrame(web_frame, text="Agregar Página Web", padding="10")
        add_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Entrada de nombre
        ttk.Label(add_frame, text="Nombre de Página:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.web_name_var = tk.StringVar()
        ttk.Entry(add_frame, textvariable=self.web_name_var, width=30).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 5), padx=(5, 0))
        
        # Entrada de URL
        ttk.Label(add_frame, text="URL Web:").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        self.web_url_var = tk.StringVar()
        url_entry = ttk.Entry(add_frame, textvariable=self.web_url_var, width=40)
        url_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 5), padx=(5, 0))
        
        # Ejemplo
        example = ttk.Label(add_frame, text="Ejemplo: https://www.youtube.com/playlist?list=...", foreground="gray")
        example.grid(row=2, column=1, sticky=tk.W, pady=(0, 10), padx=(5, 0))
        
        # Botón agregar
        ttk.Button(add_frame, text="Agregar Página Web", command=self.add_web).grid(row=3, column=0, columnspan=2)
        
        # Configurar pesos de grid
        add_frame.columnconfigure(1, weight=1)
        
        # Lista de páginas web
        list_frame = ttk.LabelFrame(web_frame, text="Páginas Web Agregadas", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Treeview para páginas web
        columns = ('name', 'url')
        self.web_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=8)
        
        self.web_tree.heading('name', text='Nombre de Página')
        self.web_tree.heading('url', text='URL')
        
        self.web_tree.column('name', width=150)
        self.web_tree.column('url', width=400)
        
        # Barra de desplazamiento
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.web_tree.yview)
        self.web_tree.configure(yscrollcommand=scrollbar.set)
        
        self.web_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Botón eliminar
        ttk.Button(list_frame, text="Eliminar Seleccionado", command=self.remove_web).pack(pady=(5, 0))

    def create_view_tab(self):
        # Marco de vista
        view_frame = ttk.Frame(self.tab_view)
        view_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Widget de texto para mostrar todos los elementos
        self.view_text = scrolledtext.ScrolledText(view_frame, wrap=tk.WORD, height=20)
        self.view_text.pack(fill=tk.BOTH, expand=True)

        self.view_text.config(state='disabled')
        
        # Nota (ya no se necesita botón de refrescar)
        ttk.Label(view_frame, text="Esta lista se actualiza automáticamente después de agregar o eliminar elementos.", 
                 font=('Arial', 9), foreground="gray").pack(pady=(5, 0))

    def create_build_tab(self):
        # Marco de construcción
        build_frame = ttk.Frame(self.tab_build)
        build_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        filename_frame = ttk.LabelFrame(build_frame, text="Configuración de Archivo Batch", padding="10")
        filename_frame.pack(fill=tk.X, pady=(0, 10))

        # --- Fila 0: Nombre de archivo ---
        ttk.Label(filename_frame, text="Nombre del Archivo Batch:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

        name_edit_frame = ttk.Frame(filename_frame)
        name_edit_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 5), padx=(5, 0))

        self.batch_filename_var = tk.StringVar(value="Abrir_Stream")
        self.filename_entry = ttk.Entry(name_edit_frame, textvariable=self.batch_filename_var, width=30)
        self.filename_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(name_edit_frame, text="Actualizar Nombre",
                command=self.update_batch_filename, width=15).pack(side=tk.RIGHT)

        # --- Fila 1: Retraso ---
        ttk.Label(filename_frame, text="Retraso Entre Aperturas de Programas (s):").grid(row=1, column=0, sticky=tk.W, pady=(5, 5))

        delay_frame = ttk.Frame(filename_frame)
        delay_frame.grid(row=1, column=1, sticky=tk.W, pady=(5, 5), padx=(5, 0))

        self.delay_var = tk.IntVar(value=2)
        self.delay_spinbox = ttk.Spinbox(delay_frame, from_=1, to=30,
                                        textvariable=self.delay_var, width=5)
        self.delay_spinbox.pack(side=tk.LEFT)
        ttk.Label(delay_frame, text="segundos", foreground="gray").pack(side=tk.LEFT, padx=(5, 0))

        # --- Fila 2: Etiqueta de ayuda ---
        ttk.Label(filename_frame,
                text="La extensión .bat se agrega automáticamente. El retraso aplica entre cada lanzamiento de programa.",
                font=('Arial', 9), foreground="gray").grid(row=2, column=0, columnspan=2, sticky=tk.W)

        filename_frame.columnconfigure(1, weight=1)
        
        # Marco de resumen
        summary_frame = ttk.LabelFrame(build_frame, text="Resumen", padding="10")
        summary_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.summary_text = scrolledtext.ScrolledText(summary_frame, wrap=tk.WORD, height=10)
        self.summary_text.pack(fill=tk.BOTH, expand=True)
        
        # Actualizar resumen
        self.update_summary()
        
        # Botón construir
        button_frame = ttk.Frame(build_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Nota: Eliminado botón "Actualizar Resumen" ya que se actualiza automáticamente
        self.build_button = ttk.Button(button_frame, text="⚙️Construir Archivo Batch⚙️", 
                                    command=self.build_batch, style="Accent.TButton")
        self.build_button.pack()
        
        # Barra de progreso
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(build_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
        
        # Etiqueta de estado
        self.build_status_label = ttk.Label(build_frame, text="")
        self.build_status_label.pack()
        

    def update_batch_filename(self):
        """Actualizar el nombre del archivo batch desde la entrada del usuario."""
        new_name = self.batch_filename_var.get().strip()
        
        if not new_name:
            messagebox.showwarning("Advertencia", "Por favor ingrese un nombre de archivo")
            return
        
        # Remover extensión .bat si el usuario la incluyó
        if new_name.lower().endswith('.bat'):
            new_name = new_name[:-4]
        
        # Validar nombre de archivo
        if not re.match(r'^[a-zA-Z0-9_\- ]+$', new_name):
            messagebox.showerror("Error", 
                                "Nombre de archivo inválido. Use sólo letras, números, espacios, guiones y guiones bajos.")
            # Restablecer al nombre válido anterior
            self.batch_filename_var.set("Abrir_Stream")
            return
        
        # Actualizar variable
        self.batch_filename_var.set(new_name)
        self.update_status(f"Nombre de archivo batch actualizado a: {new_name}.bat")

    def create_status_bar(self):
        self.status_bar = ttk.Label(self.main_frame, text="Listo", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

    # Métodos auxiliares
    def update_status(self, message):
        self.status_bar.config(text=message)
        self.root.update_idletasks()

    def update_obs_status(self):
        if self.obs_path:
            self.obs_status_label.config(text=f"✓ {self.obs_path}", style='Success.TLabel')
            # Mostrar botón limpiar cuando OBS está configurado
            self.clear_obs_button.pack(side=tk.RIGHT)
        else:
            self.obs_status_label.config(text="No hay OBS configurado", style='Error.TLabel')
            # Ocultar botón limpiar cuando no hay OBS configurado
            self.clear_obs_button.pack_forget()

    def refresh_view(self):
        # Habilitar widget de texto para edición
        self.view_text.config(state='normal')
        self.view_text.delete(1.0, tk.END)
        
        # Agregar OBS
        self.view_text.insert(tk.END, "📷 SOFTWARE DE TRANSMISIÓN:\n")
        if self.obs_path:
            self.view_text.insert(tk.END, f"  • OBS: {self.obs_path}\n")
        else:
            self.view_text.insert(tk.END, "  • No hay OBS configurado\n")
        self.view_text.insert(tk.END, "\n")
        
        # Agregar juegos Steam
        self.view_text.insert(tk.END, "🎮 JUEGOS/PROGRAMAS DE STEAM:\n")
        if self.steam_dictionary:
            for name, url in self.steam_dictionary.items():
                self.view_text.insert(tk.END, f"  • {name}: {url}\n")
        else:
            self.view_text.insert(tk.END, "  • No hay juegos/programas de Steam agregados\n")
        self.view_text.insert(tk.END, "\n")
        
        # Agregar Programas
        self.view_text.insert(tk.END, "💻 OTROS PROGRAMAS:\n")
        if self.program_shortcut_dictionary:
            for name, path in self.program_shortcut_dictionary.items():
                self.view_text.insert(tk.END, f"  • {name}: {path}\n")
        else:
            self.view_text.insert(tk.END, "  • No hay programas agregados\n")
        self.view_text.insert(tk.END, "\n")
        
        # Agregar apps de Tienda MS
        self.view_text.insert(tk.END, "🏪 APPS DE LA TIENDA MICROSOFT:\n")
        if self.ms_store_dictionary:
            for name, aumid in self.ms_store_dictionary.items():
                self.view_text.insert(tk.END, f"  • {name}: {aumid}\n")
        else:
            self.view_text.insert(tk.END, "  • No hay apps de la Tienda MS agregadas\n")
        self.view_text.insert(tk.END, "\n")

        # Agregar páginas web
        self.view_text.insert(tk.END, "🌐 PÁGINAS WEB:\n")
        if self.web_dictionary:
            for name, url in self.web_dictionary.items():
                self.view_text.insert(tk.END, f"  • {name}: {url}\n")
        else:
            self.view_text.insert(tk.END, "  • No hay páginas web agregadas\n")
        
        # Deshabilitar widget de texto para prevenir edición
        self.view_text.config(state='disabled')

    def update_summary(self):
        self.summary_text.config(state='normal')
        self.summary_text.delete(1.0, tk.END)
        
        total_items = 0
        if self.obs_path:
            total_items += 1
        
        total_items += len(self.steam_dictionary)
        total_items += len(self.program_shortcut_dictionary)
        total_items += len(self.web_dictionary)
        total_items += len(self.ms_store_dictionary)

        self.summary_text.insert(tk.END, f"LISTA DE LANZAMIENTO\n")
        self.summary_text.insert(tk.END, f"{'='*40}\n\n")
        self.summary_text.insert(tk.END, f"Total de elementos en lista de lanzamiento: {total_items}\n\n")

        self.summary_text.insert(tk.END, f"OBS: {'✓ Configurado' if self.obs_path else 'X No configurado'}\n")
        self.summary_text.insert(tk.END, f"Juegos/Programas de Steam: {len(self.steam_dictionary)}\n")
        self.summary_text.insert(tk.END, f"Programas: {len(self.program_shortcut_dictionary)}\n")
        self.summary_text.insert(tk.END, f"Apps de Tienda MS: {len(self.ms_store_dictionary)}\n")
        self.summary_text.insert(tk.END, f"Páginas Web: {len(self.web_dictionary)}\n\n")
        
        if total_items > 0:
            self.summary_text.insert(tk.END, "¡Listo para construir archivo batch!\n")
        else:
            self.summary_text.insert(tk.END, "¡ Por favor agregue al menos un programa/juego/página antes de construir el archivo batch !\n")

        self.summary_text.config(state='disabled')

    # NUEVO: Método auto-refrescar
    def auto_refresh(self):
        """Refrescar automáticamente la lista de vista y resumen de construcción."""
        self.refresh_view()
        self.update_summary()

    # NUEVO: Limpiar ruta removiendo comillas
    def clean_path(self, path):
        """Remover comillas y espacios en blanco de una ruta."""
        if not path:
            return path
        # Remover comillas simples y dobles
        cleaned = path.replace('"', '').replace("'", "")
        # Remover espacios en blanco
        cleaned = cleaned.strip()
        return cleaned

    # Métodos de navegación de archivos
    def browse_obs(self):
        filename = filedialog.askopenfilename(
            title="Seleccionar Ejecutable de OBS",
            filetypes=[("Archivos ejecutables", "*.exe"), ("Todos los archivos", "*.*")]
        )
        if filename:
            # Limpiar ruta
            cleaned_path = self.clean_path(filename)
            path = Path(cleaned_path)
            if path.exists() and path.suffix.lower() == '.exe':
                self.obs_path_var.set(str(path))
                self.obs_path = str(path)
                self.update_obs_status()
                self.update_status("Ruta de OBS actualizada exitosamente")
                # Auto-refrescar después de agregar
                self.auto_refresh()
            else:
                messagebox.showerror("Error", "Por favor seleccione un archivo .exe válido")

    # NUEVO: Establecer ruta OBS desde entrada de texto
    def set_obs_path(self):
        path_input = self.obs_path_var.get().strip()
        
        if not path_input:
            messagebox.showwarning("Advertencia", "Por favor ingrese una ruta para OBS")
            return
        
        # Limpiar ruta (remover comillas)
        cleaned_path = self.clean_path(path_input)
        
        try:
            path = Path(cleaned_path)
            
            # Verificar si la ruta existe
            if not path.exists():
                messagebox.showerror("Error", f"¡La ruta '{cleaned_path}' no existe!")
                return
            
            # Verificar si es un archivo
            if not path.is_file():
                messagebox.showerror("Error", f"¡'{cleaned_path}' no es un archivo!")
                return
            
            # Verificar si es un archivo .exe
            if path.suffix.lower() != '.exe':
                messagebox.showerror("Error", f"¡'{cleaned_path}' no es un archivo .exe!")
                return
            
            # Establecer ruta OBS
            self.obs_path = str(path)
            self.update_obs_status()
            self.update_status("OBS establecido exitosamente desde entrada manual de ruta")
            # Auto-refrescar después de agregar
            self.auto_refresh()
            
        except Exception as e:
            messagebox.showerror("Error", f"Ruta inválida: {str(e)}")

    def browse_program(self):
        filename = filedialog.askopenfilename(
            title="Seleccionar Ejecutable de Programa",
            filetypes=[("Archivos ejecutables", "*.exe"), ("Todos los archivos", "*.*")]
        )
        if filename:
            # Limpiar ruta
            cleaned_path = self.clean_path(filename)
            path = Path(cleaned_path)
            if path.exists() and path.suffix.lower() == '.exe':
                self.prog_path_var.set(str(path))
            else:
                messagebox.showerror("Error", "Por favor seleccione un archivo .exe válido")

    # Métodos agregar (actualizados para limpiar rutas y auto-refrescar)
    def add_program(self):
        name = self.prog_name_var.get().strip()
        path_input = self.prog_path_var.get().strip()
        
        if not name:
            messagebox.showwarning("Advertencia", "Por favor ingrese un nombre de programa")
            return
        
        if not path_input:
            messagebox.showwarning("Advertencia", "Por favor ingrese una ruta de ejecutable")
            return
        
        # Limpiar ruta (remover comillas)
        cleaned_path = self.clean_path(path_input)
        
        try:
            path_obj = Path(cleaned_path)
            
            if not path_obj.exists():
                messagebox.showerror("Error", f"¡La ruta '{cleaned_path}' no existe!")
                return
            
            if not path_obj.is_file():
                messagebox.showerror("Error", f"¡'{cleaned_path}' no es un archivo!")
                return
            
            if path_obj.suffix.lower() != '.exe':
                messagebox.showerror("Error", f"¡'{cleaned_path}' no es un archivo .exe!")
                return
            
            # Guardar ruta limpia
            self.program_shortcut_dictionary[name] = str(path_obj)
            
            # Actualizar treeview
            self.programs_tree.insert('', tk.END, values=(name, str(path_obj)))
            
            # Limpiar entradas
            self.prog_name_var.set('')
            self.prog_path_var.set('')
            
            self.update_status(f"Programa agregado: {name}")
            # Auto-refrescar después de agregar
            self.auto_refresh()
            
        except Exception as e:
            messagebox.showerror("Error", f"Ruta inválida: {str(e)}")

    def add_steam(self):
        name = self.steam_name_var.get().strip()
        url = self.steam_url_var.get().strip()
        
        if not name:
            messagebox.showwarning("Advertencia", "Por favor ingrese un nombre de juego/programa de Steam")
            return
        
        if not url:
            messagebox.showwarning("Advertencia", "Por favor ingrese una URL de Steam")
            return
        
        # Limpiar URL (remover comillas)
        cleaned_url = self.clean_path(url)
        
        is_valid, message = validate_steam_url(cleaned_url)
        
        if not is_valid:
            messagebox.showerror("Error", f"URL de Steam inválida: {message}")
            return
        
        self.steam_dictionary[name] = cleaned_url
        
        # Actualizar treeview
        self.steam_tree.insert('', tk.END, values=(name, cleaned_url))
        
        # Limpiar entradas
        self.steam_name_var.set('')
        self.steam_url_var.set('')
        
        self.update_status(f"Juego/programa de Steam agregado: {name}")
        # Auto-refrescar después de agregar
        self.auto_refresh()

    def search_ms_store_apps(self):
        query = self.ms_search_var.get().strip()
        if not query:
            messagebox.showwarning("Advertencia", "Por favor ingrese un nombre de app para buscar")
            return

        self.ms_search_button.config(state='disabled')
        self.ms_search_status.config(text="Buscando...", foreground="gray")
        self.ms_results_tree.delete(*self.ms_results_tree.get_children())

        thread = threading.Thread(target=self._search_ms_store_thread, args=(query,), daemon=True)
        thread.start()

    def _search_ms_store_thread(self, query):
        import subprocess, json
        ps_cmd = (
            f'Get-StartApps | Where-Object {{$_.Name -like "*{query}*"}} '
            f'| Select-Object Name, AppID | ConvertTo-Json'
        )
        try:
            result = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_cmd],
                capture_output=True, text=True, timeout=15
            )
            raw = result.stdout.strip()
            if not raw:
                self.root.after(0, lambda: self._ms_search_done([]))
                return

            data = json.loads(raw)
            if isinstance(data, dict):
                data = [data]
            apps = [(item['Name'], item['AppID']) for item in data if item.get('AppID')]
            self.root.after(0, lambda: self._ms_search_done(apps))
        except Exception as e:
            err = str(e)
            self.root.after(0, lambda: self._ms_search_error(err))

    def _ms_search_done(self, apps):
        self.ms_search_button.config(state='normal')
        if not apps:
            self.ms_search_status.config(text="No se encontraron apps. Intente con otro nombre.", foreground="red")
            return
        self.ms_search_status.config(text=f"{len(apps)} resultado(s) encontrado(s). Seleccione uno y haga clic en 'Agregar App Seleccionada'.", foreground="green")
        for name, aumid in apps:
            self.ms_results_tree.insert('', tk.END, values=(name, aumid))

    def _ms_search_error(self, error):
        self.ms_search_button.config(state='normal')
        self.ms_search_status.config(text=f"Búsqueda fallida: {error}", foreground="red")

    def add_ms_store_from_results(self):
        selected = self.ms_results_tree.selection()
        if not selected:
            messagebox.showwarning("Advertencia", "Por favor seleccione una app de los resultados de búsqueda")
            return

        values = self.ms_results_tree.item(selected[0], 'values')
        name, aumid = values[0], values[1]

        if name in self.ms_store_dictionary:
            messagebox.showinfo("Información", f"'{name}' ya está en la lista de lanzamiento.")
            return

        self.ms_store_dictionary[name] = aumid
        self.ms_store_tree.insert('', tk.END, values=(name, aumid))
        self.update_status(f"App de Tienda MS agregada: {name}")
        self.auto_refresh()

    def add_web(self):
        name = self.web_name_var.get().strip()
        url = self.web_url_var.get().strip()
        
        if not name:
            messagebox.showwarning("Advertencia", "Por favor ingrese un nombre de página")
            return
        
        if not url:
            messagebox.showwarning("Advertencia", "Por favor ingrese una URL web")
            return
        
        # Limpiar URL (remover comillas)
        cleaned_url = self.clean_path(url)
        
        if not cleaned_url.startswith(('http://', 'https://')):
            cleaned_url = 'https://' + cleaned_url
        
        self.web_dictionary[name] = cleaned_url
        
        # Actualizar treeview
        self.web_tree.insert('', tk.END, values=(name, cleaned_url))
        
        # Limpiar entradas
        self.web_name_var.set('')
        self.web_url_var.set('')
        
        self.update_status(f"Página web agregada: {name}")
        # Auto-refrescar después de agregar
        self.auto_refresh()

    # Métodos eliminar (actualizados para auto-refrescar)
    def remove_program(self):
        selected = self.programs_tree.selection()
        if not selected:
            messagebox.showwarning("Advertencia", "Por favor seleccione un programa para eliminar")
            return
        
        for item in selected:
            values = self.programs_tree.item(item, 'values')
            name = values[0]
            if name in self.program_shortcut_dictionary:
                del self.program_shortcut_dictionary[name]
            self.programs_tree.delete(item)
        
        self.update_status("Programa eliminado")
        # Auto-refrescar después de eliminar
        self.auto_refresh()

    def remove_steam(self):
        selected = self.steam_tree.selection()
        if not selected:
            messagebox.showwarning("Advertencia", "Por favor seleccione un juego/programa de Steam para eliminar")
            return
        
        for item in selected:
            values = self.steam_tree.item(item, 'values')
            name = values[0]
            if name in self.steam_dictionary:
                del self.steam_dictionary[name]
            self.steam_tree.delete(item)
        
        self.update_status("Juego/programa de Steam eliminado")
        # Auto-refrescar después de eliminar
        self.auto_refresh()

    def remove_ms_store(self):
        selected = self.ms_store_tree.selection()
        if not selected:
            messagebox.showwarning("Advertencia", "Por favor seleccione una app para eliminar")
            return

        for item in selected:
            values = self.ms_store_tree.item(item, 'values')
            name = values[0]
            if name in self.ms_store_dictionary:
                del self.ms_store_dictionary[name]
            self.ms_store_tree.delete(item)

        self.update_status("App de Tienda MS eliminada")
        self.auto_refresh()

    def remove_web(self):
        selected = self.web_tree.selection()
        if not selected:
            messagebox.showwarning("Advertencia", "Por favor seleccione una página web para eliminar")
            return
        
        for item in selected:
            values = self.web_tree.item(item, 'values')
            name = values[0]
            if name in self.web_dictionary:
                del self.web_dictionary[name]
            self.web_tree.delete(item)
        
        self.update_status("Página web eliminada")
        # Auto-refrescar después de eliminar
        self.auto_refresh()

    # Método construir
    def build_batch(self):
        # Verificar si hay algo para construir
        if not self.obs_path and not self.steam_dictionary and not self.program_shortcut_dictionary and not self.web_dictionary and not self.ms_store_dictionary:
            messagebox.showwarning("Advertencia", "Por favor agregue al menos un elemento (programa/juego/página) antes de construir")
            return
        
        # Confirmar con usuario
        response = messagebox.askyesno("Confirmar Construcción", 
                                      "Esto creará un archivo batch (.bat) que lanzará todos los elementos agregados.\n\n"
                                      "¿Desea proceder?")
        if not response:
            return
        
        # Deshabilitar botón construir durante el proceso
        self.build_button.config(state='disabled')
        self.progress_var.set(0)
        self.build_status_label.config(text="Construyendo archivo .bat...")
        
        # Ejecutar en hilo separado para mantener GUI responsiva
        thread = threading.Thread(target=self._build_batch_thread)
        thread.daemon = True
        thread.start()

    def _build_batch_thread(self):
        # Inicializar COM para este hilo
        import pythoncom
        pythoncom.CoInitialize()
        
        try:
            # Paso 1: Crear carpeta de accesos directos
            self.root.after(0, lambda: self.update_status("Creando carpeta de accesos directos..."))
            self.root.after(0, lambda: self.progress_var.set(20))
            
            shortcut_folder_dir = make_empty_shortcut_folder()
            
            # Paso 2: Llenar carpeta de accesos directos y obtener ruta de acceso directo OBS
            self.root.after(0, lambda: self.update_status("Creando accesos directos..."))
            self.root.after(0, lambda: self.progress_var.set(40))
            
            obs_shortcut_path = fill_shortcut_folder(
                shortcut_dir=shortcut_folder_dir,
                S_dictionary=self.program_shortcut_dictionary if self.program_shortcut_dictionary else None,
                steam_dict=self.steam_dictionary if self.steam_dictionary else None,
                obs_dir=self.obs_path,
                web_dict=self.web_dictionary if self.web_dictionary else None,
                ms_store_dict=self.ms_store_dictionary if self.ms_store_dictionary else None
            )

            # Paso 3: Crear archivo batch con ruta de acceso directo OBS
            self.root.after(0, lambda: self.update_status("Creando archivo batch..."))
            self.root.after(0, lambda: self.progress_var.set(60))

            batch_filename = self.batch_filename_var.get() + ".bat"
            delay_seconds = self.delay_var.get()

            create_bat_file(
                self.program_shortcut_dictionary,
                shortcut_folder_dir,
                obs_short_dir=obs_shortcut_path,
                batch_filename=batch_filename,
                delay_seconds=delay_seconds
            )
            
            # Paso 4: Completar
            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(0, lambda: self.update_status("¡Archivo batch creado exitosamente!"))
            self.root.after(0, lambda: self.build_status_label.config(text="✓ ¡Archivo batch creado exitosamente!"))
            
            # Mostrar mensaje de éxito
            self.root.after(0, lambda: messagebox.showinfo("Éxito", 
                f"¡Archivo batch creado exitosamente!\n\n"
                f"El archivo '{batch_filename}' ha sido creado en el directorio actual.\n"
                f"¡Haga doble clic en él para lanzar todos sus programas!\n\n"
                f"IMPORTANTE: No mueva la carpeta de 'Accesos_para_bat', o el archivo batch no funcionará."))
            
        except Exception as e:
            # Capturar excepción en el ámbito local del lambda
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self.update_status(f"Error: {msg}"))
            self.root.after(0, lambda msg=error_msg: self.build_status_label.config(text=f"X Error: {msg}", style='Error.TLabel'))
            self.root.after(0, lambda msg=error_msg: messagebox.showerror("Error", f"Ocurrió un error:\n{msg}"))
        finally:
            # Desinicializar COM
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self.build_button.config(state='normal'))

# Mantener funciones originales (deben permanecer sin cambios)
def make_empty_shortcut_folder():
    current_folder = os.getcwd()
    shortcut_folder_path = f"{current_folder}\\Accesos_para_bat"
    obs_shortcut_path = f"{shortcut_folder_path}\\OBS"

    if os.path.exists(shortcut_folder_path):
        shutil.rmtree(shortcut_folder_path)

    os.makedirs(shortcut_folder_path, exist_ok=True)

    return shortcut_folder_path

def create_exe_shortcut(exe_path, shortcut_dir):
    exe_path = Path(exe_path)
    shortcut_dir = Path(shortcut_dir)
    shortcut_dir.mkdir(parents=True, exist_ok=True)

    name = exe_path.stem

    shortcut_path = shortcut_dir / f"{name}.lnk"

    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(str(shortcut_path))

    shortcut.TargetPath = str(exe_path)
    shortcut.WorkingDirectory = str(exe_path.parent)
    shortcut.IconLocation = str(exe_path)
    shortcut.Save()

    return shortcut_path

def fill_shortcut_folder(shortcut_dir, S_dictionary=None, steam_dict=None, obs_dir=None, web_dict=None, ms_store_dict=None):
    obs_shortcut = None
    if obs_dir is not None:
        obs_shortcut_path = f"{shortcut_dir}\\OBS"
        create_exe_shortcut(obs_dir, obs_shortcut_path)
        obs_shortcut = obs_shortcut_path
        time.sleep(1)

    if S_dictionary is not None:
        time.sleep(1)
        for ProgName, path in S_dictionary.items():
            create_exe_shortcut(path, shortcut_dir)

    if steam_dict is not None:
        time.sleep(1)
        for game_name, app_url in steam_dict.items():
            create_steam_shortcut(game_name, app_url, shortcut_dir)

    if ms_store_dict is not None:
        time.sleep(1)
        for app_name, aumid in ms_store_dict.items():
            create_ms_store_shortcut(app_name, aumid, shortcut_dir)

    if web_dict is not None:
        time.sleep(1)
        for web_name, web_url in web_dict.items():
            create_web_shortcut(web_name, web_url, shortcut_dir)
    return obs_shortcut

def create_bat_file(S_dictionary, shortcut_dir, obs_short_dir=None, batch_filename="Abrir_Stream.bat", delay_seconds=2):
    dir_line = f'for %%a in ("{shortcut_dir}\\*.lnk") do (\n    start "" "%%~fa"\n    timeout /t {delay_seconds} /nobreak > nul\n)\n'
    start_lines = ["@echo off\n", dir_line]
    
    if obs_short_dir is not None:
        obs_line = f'for %%a in ("{obs_short_dir}\\*.lnk") do (\n    start \"\" /wait \"%%~fa\"\n)\n'
        start_lines.append(obs_line)
    
    start_lines.append("exit")

    with open(batch_filename, "w") as batfile:
        batfile.writelines(start_lines)

def create_steam_shortcut(game_name, steam_url, shortcut_dir):
    shortcut_dir = Path(shortcut_dir)
    shortcut_dir.mkdir(parents=True, exist_ok=True)

    shortcut_path = shortcut_dir / f"{game_name}.lnk"

    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(str(shortcut_path))
    
    shortcut.TargetPath = steam_url
    
    shortcut.Save()

    return shortcut_path

def create_web_shortcut(web_name, web_url, shortcut_dir):
    shortcut_dir = Path(shortcut_dir)
    shortcut_dir.mkdir(parents=True, exist_ok=True)

    shortcut_path = shortcut_dir / f"{web_name}.lnk"

    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(str(shortcut_path))
    
    shortcut.TargetPath = web_url
    
    shortcut.Save()

    return shortcut_path

def create_ms_store_shortcut(app_name, aumid, shortcut_dir):
    shortcut_dir = Path(shortcut_dir)
    shortcut_dir.mkdir(parents=True, exist_ok=True)

    shortcut_path = shortcut_dir / f"{app_name}.lnk"

    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(str(shortcut_path))

    shortcut.TargetPath = r"C:\Windows\explorer.exe"
    shortcut.Arguments = f"shell:AppsFolder\\{aumid}"

    shortcut.Save()

    return shortcut_path

def validate_steam_url(url):
    url = url.strip()
    
    if url.startswith('steam://'):
        steam_patterns = [
            r'^steam://rungameid/\d+$',
            r'^steam://run/\d+$', 
            r'^steam://launch/\d+$',
            r'^steam://openurl/[^ ]+$'
        ]
        
        for pattern in steam_patterns:
            if re.match(pattern, url):
                return True, "URL de Steam válida."
        
        return False, "Formato de URL de Steam inválido. Debe verse de esta manera: steam://rungameid/123456."
    
    else:
        return False, "La URL de Steam debe comenzar con 'steam://'."

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def main():
    root = tk.Tk()

    icon_path = resource_path("Batpic_Icon.ico")

    try:
        root.iconbitmap(icon_path)
    except Exception as e:
        print(f"No se pudo cargar el ícono: {e}")
    
    app = StreamBatGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()