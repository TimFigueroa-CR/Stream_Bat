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
        
        # Initialize program dictionaries and OBS path/shortcut
        self.obs_path = None
        self.steam_dictionary = {}
        self.program_shortcut_dictionary = {}
        self.web_dictionary = {}
        self.ms_store_dictionary = {}
        self.obs_shortcut = None
        
        #Setup
        self.setup_styles()
        
        # Create main frame
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)
        
        # Create header, notebook(for the tabs), status bar and center the window on the screen
        self.create_header()
        
        self.create_notebook()

        self.create_status_bar()
        
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
                            text="Create a batch(.bat) file to launch all your programs with one click!")
        subtitle.pack(side=tk.LEFT, padx=(10, 0))

        author_label = ttk.Label(header_frame, text="By: Timothy Figueroa, with\nhelp from Deepseek-V3", font=('Arial', 8), foreground="gray")
        author_label.pack(side=tk.RIGHT)

    def create_notebook(self):
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Create tabs
        self.tab_obs = ttk.Frame(self.notebook)
        self.tab_programs = ttk.Frame(self.notebook)
        self.tab_steam = ttk.Frame(self.notebook)
        self.tab_ms_store = ttk.Frame(self.notebook)
        self.tab_web = ttk.Frame(self.notebook)
        self.tab_view = ttk.Frame(self.notebook)
        self.tab_build = ttk.Frame(self.notebook)

        # Add each tab to the notebook
        self.notebook.add(self.tab_obs, text="📷 OBS")
        self.notebook.add(self.tab_programs, text="💻 Programs")
        self.notebook.add(self.tab_steam, text="🎮 Steam")
        self.notebook.add(self.tab_ms_store, text="🏪 MS Store")
        self.notebook.add(self.tab_web, text="🌐 Web Pages")
        self.notebook.add(self.tab_view, text="👁️ View List")
        self.notebook.add(self.tab_build, text="⚙️ Build")

        # Populate each tab accordingly
        self.create_obs_tab()
        self.create_programs_tab()
        self.create_steam_tab()
        self.create_ms_store_tab()
        self.create_web_tab()
        self.create_view_tab()
        self.create_build_tab()

    def create_obs_tab(self):
        # OBS tab/frame
        obs_frame = ttk.LabelFrame(self.tab_obs, text="OBS Studio Settings", padding="10")
        obs_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        ttk.Label(obs_frame, text="Add your OBS Studio executable:").pack(anchor=tk.W, pady=(0, 10))
        
        #Instructions for the file path
        ttk.Label(obs_frame, text="You can either browse for the file or paste the path manually:").pack(anchor=tk.W, pady=(0, 5))
        
        # Frame to insert the path
        path_frame = ttk.Frame(obs_frame)
        path_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.obs_path_var = tk.StringVar()
        self.obs_path_entry = ttk.Entry(path_frame, textvariable=self.obs_path_var, width=50)
        self.obs_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        # 'Set Path' and 'Browse' button frames
        button_frame = ttk.Frame(path_frame)
        button_frame.pack(side=tk.RIGHT)
        
        ttk.Button(button_frame, text="Browse...", command=self.browse_obs, width=10).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Set Path", command=self.set_obs_path, width=10).pack(side=tk.RIGHT)
        
        # OBS status(path added or no path)
        self.obs_status_frame = ttk.LabelFrame(obs_frame, text="Current OBS", padding="10")
        self.obs_status_frame.pack(fill=tk.X, pady=(10, 0))
        
        # OBS Status and 'clear' button frame(below the path frame)
        status_content_frame = ttk.Frame(self.obs_status_frame)
        status_content_frame.pack(fill=tk.X, expand=True)
        
        self.obs_status_label = ttk.Label(status_content_frame, text="No OBS configured")
        self.obs_status_label.pack(side=tk.LEFT, anchor=tk.W, fill=tk.X, expand=True)
        
        self.clear_obs_button = ttk.Button(status_content_frame, text="Clear OBS", 
                                           command=self.clear_obs_path, width=10)
        self.clear_obs_button.pack(side=tk.RIGHT)
        
        # Make sure 'Clear OBS' button is only displayed when there is an actual path entered
        self.clear_obs_button.pack_forget()
        
        # Update
        self.update_obs_status()

    def clear_obs_path(self):
        """Clear/reset OBS path"""
        if self.obs_path:
            # Confirmation Box
            response = messagebox.askyesno("Clear OBS", 
                                            "Are you sure you want to clear the OBS path?\n\n"
                                            "This will remove OBS from your launch list.")
            if response:
                #If yes, remove the path
                self.obs_path = None
                self.obs_path_var.set('')

                # Update the status
                self.update_obs_status()
                
                # Refresh to update the view list
                self.auto_refresh()
                
                self.update_status("OBS path cleared")
        else:
            messagebox.showinfo("Info", "No OBS path is currently configured.")

    def create_programs_tab(self):
        # Frame for non-Steam programs
        programs_frame = ttk.Frame(self.tab_programs)
        programs_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
        add_frame = ttk.LabelFrame(programs_frame, text="Add Program", padding="10")
        add_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Entry for program name and path

        ttk.Label(add_frame, text="Program Name:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.prog_name_var = tk.StringVar()
        ttk.Entry(add_frame, textvariable=self.prog_name_var, width=30).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 5), padx=(5, 0))
        
        ttk.Label(add_frame, text="Executable Path:").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        path_frame = ttk.Frame(add_frame)
        path_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 5), padx=(5, 0))
        
        self.prog_path_var = tk.StringVar()
        self.prog_path_entry = ttk.Entry(path_frame, textvariable=self.prog_path_var, width=40)
        self.prog_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        ttk.Button(path_frame, text="Browse...", command=self.browse_program).pack(side=tk.RIGHT)
        
        # 'Add Program' button
        ttk.Button(add_frame, text="Add Program", command=self.add_program).grid(row=2, column=0, columnspan=2, pady=(10, 0))
        
        # Configure grid
        add_frame.columnconfigure(1, weight=1)
        
        # Frame/display programs that have been added
        list_frame = ttk.LabelFrame(programs_frame, text="Added Programs", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ('name', 'path')
        self.programs_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=8)
        
        self.programs_tree.heading('name', text='Program Name')
        self.programs_tree.heading('path', text='Executable Path')
        
        self.programs_tree.column('name', width=150)
        self.programs_tree.column('path', width=400)
        
        # Scrollbar for program list
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.programs_tree.yview)
        self.programs_tree.configure(yscrollcommand=scrollbar.set)
        
        self.programs_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 'Remove program' button
        ttk.Button(list_frame, text="Remove Selected", command=self.remove_program).pack(pady=(5, 0))

    def create_steam_tab(self):
        # Frame for Steam programs
        steam_frame = ttk.Frame(self.tab_steam)
        steam_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        add_frame = ttk.LabelFrame(steam_frame, text="Add Steam Game/Program", padding="10")
        add_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Entry for Steam program name and URL
        ttk.Label(add_frame, text="Game/Program Name:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.steam_name_var = tk.StringVar()
        ttk.Entry(add_frame, textvariable=self.steam_name_var, width=30).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 5), padx=(5, 0))
        
        ttk.Label(add_frame, text="Steam URL:").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        self.steam_url_var = tk.StringVar()
        url_entry = ttk.Entry(add_frame, textvariable=self.steam_url_var, width=40)
        url_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 5), padx=(5, 0))
        
        # URL example to display to the user under the entry boxes
        example = ttk.Label(add_frame, text="Example: steam://rungameid/123456", foreground="gray")
        example.grid(row=2, column=1, sticky=tk.W, pady=(0, 10), padx=(5, 0))
        
        # 'Add Steam program' button
        ttk.Button(add_frame, text="Add Steam Game/Program", command=self.add_steam).grid(row=3, column=0, columnspan=2)
        
        # Configure grid
        add_frame.columnconfigure(1, weight=1)
        
        # Add/display a list for steam games/programs that have been added
        list_frame = ttk.LabelFrame(steam_frame, text="Added Steam Games/Programs", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ('name', 'url')
        self.steam_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=8)
        
        self.steam_tree.heading('name', text='Game/Program Name')
        self.steam_tree.heading('url', text='Steam URL')
        
        self.steam_tree.column('name', width=150)
        self.steam_tree.column('url', width=400)
        
        # Scrollbar for the Steam program list
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.steam_tree.yview)
        self.steam_tree.configure(yscrollcommand=scrollbar.set)
        
        self.steam_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 'Remove Steam Program' button
        ttk.Button(list_frame, text="Remove Selected", command=self.remove_steam).pack(pady=(5, 0))

    def create_ms_store_tab(self):
        ms_frame = ttk.Frame(self.tab_ms_store)
        ms_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Search section
        search_frame = ttk.LabelFrame(ms_frame, text="Search for a Microsoft Store App", padding="10")
        search_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(search_frame, text="Type the app name and click Search:").pack(anchor=tk.W, pady=(0, 5))

        input_row = ttk.Frame(search_frame)
        input_row.pack(fill=tk.X)

        self.ms_search_var = tk.StringVar()
        self.ms_search_entry = ttk.Entry(input_row, textvariable=self.ms_search_var, width=35)
        self.ms_search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.ms_search_entry.bind('<Return>', lambda e: self.search_ms_store_apps())

        self.ms_search_button = ttk.Button(input_row, text="Search", command=self.search_ms_store_apps, width=10)
        self.ms_search_button.pack(side=tk.LEFT)

        self.ms_search_status = ttk.Label(search_frame, text="", foreground="gray", font=('Arial', 9))
        self.ms_search_status.pack(anchor=tk.W, pady=(5, 0))

        # Results section
        results_frame = ttk.LabelFrame(ms_frame, text="Search Results — select an app and click Add", padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        res_columns = ('name', 'aumid')
        self.ms_results_tree = ttk.Treeview(results_frame, columns=res_columns, show='headings', height=5)
        self.ms_results_tree.heading('name', text='App Name')
        self.ms_results_tree.heading('aumid', text='App ID (AUMID)')
        self.ms_results_tree.column('name', width=160)
        self.ms_results_tree.column('aumid', width=370)

        res_scroll = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.ms_results_tree.yview)
        self.ms_results_tree.configure(yscrollcommand=res_scroll.set)
        self.ms_results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        res_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Button(results_frame, text="Add Selected App", command=self.add_ms_store_from_results).pack(pady=(5, 0))

        # Added apps list
        list_frame = ttk.LabelFrame(ms_frame, text="Added Microsoft Store Apps", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True)

        columns = ('name', 'aumid')
        self.ms_store_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=4)
        self.ms_store_tree.heading('name', text='App Name')
        self.ms_store_tree.heading('aumid', text='App ID (AUMID)')
        self.ms_store_tree.column('name', width=160)
        self.ms_store_tree.column('aumid', width=370)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.ms_store_tree.yview)
        self.ms_store_tree.configure(yscrollcommand=scrollbar.set)
        self.ms_store_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Button(list_frame, text="Remove Selected", command=self.remove_ms_store).pack(pady=(5, 0))

    def create_web_tab(self):
        # Frame for web links/URLs
        web_frame = ttk.Frame(self.tab_web)
        web_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        add_frame = ttk.LabelFrame(web_frame, text="Add Web Page", padding="10")
        add_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Entries for the name and URL of the website
        ttk.Label(add_frame, text="Page Name:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.web_name_var = tk.StringVar()
        ttk.Entry(add_frame, textvariable=self.web_name_var, width=30).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 5), padx=(5, 0))
        
        ttk.Label(add_frame, text="Web URL:").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        self.web_url_var = tk.StringVar()
        url_entry = ttk.Entry(add_frame, textvariable=self.web_url_var, width=40)
        url_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 5), padx=(5, 0))
        
        # URL example to display to the user under the entry boxes
        example = ttk.Label(add_frame, text="Example: https://www.youtube.com/playlist?list=...", foreground="gray")
        example.grid(row=2, column=1, sticky=tk.W, pady=(0, 10), padx=(5, 0))
        
        # 'Add Web Page' button
        ttk.Button(add_frame, text="Add Web Page", command=self.add_web).grid(row=3, column=0, columnspan=2)
        
        # Configure grid weights
        add_frame.columnconfigure(1, weight=1)
        
        # Add/display a list for Web URLs that have been added
        list_frame = ttk.LabelFrame(web_frame, text="Added Web Pages", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ('name', 'url')
        self.web_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=8)
        
        self.web_tree.heading('name', text='Page Name')
        self.web_tree.heading('url', text='URL')
        
        self.web_tree.column('name', width=150)
        self.web_tree.column('url', width=400)
        
        # Scrollbar for the added URLs
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.web_tree.yview)
        self.web_tree.configure(yscrollcommand=scrollbar.set)
        
        self.web_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 'Remove URL' button
        ttk.Button(list_frame, text="Remove Selected", command=self.remove_web).pack(pady=(5, 0))

    def create_view_tab(self):
        # Frame for the View List tab
        view_frame = ttk.Frame(self.tab_view)
        view_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Display the launch list in a text widget
        self.view_text = scrolledtext.ScrolledText(view_frame, wrap=tk.WORD, height=20)
        self.view_text.pack(fill=tk.BOTH, expand=True)

        self.view_text.config(state='disabled')
        
        #Bottom label that clarifies that the list is (basically) almost updated
        ttk.Label(view_frame, text="This list updates automatically after adding or removing items.", 
                 font=('Arial', 9), foreground="gray").pack(pady=(5, 0))

    def create_build_tab(self):
        build_frame = ttk.Frame(self.tab_build)
        build_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        filename_frame = ttk.LabelFrame(build_frame, text="Batch File Configuration", padding="10")
        filename_frame.pack(fill=tk.X, pady=(0, 10))

        # --- Row 0: File name ---
        ttk.Label(filename_frame, text="Batch File Name:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

        name_edit_frame = ttk.Frame(filename_frame)
        name_edit_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 5), padx=(5, 0))

        self.batch_filename_var = tk.StringVar(value="Open_Stream")
        self.filename_entry = ttk.Entry(name_edit_frame, textvariable=self.batch_filename_var, width=30)
        self.filename_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(name_edit_frame, text="Update Name",
                command=self.update_batch_filename, width=12).pack(side=tk.RIGHT)

        # --- Row 1: Delay ---
        ttk.Label(filename_frame, text="Delay Between Programs Opening (s):").grid(row=1, column=0, sticky=tk.W, pady=(5, 5))

        delay_frame = ttk.Frame(filename_frame)
        delay_frame.grid(row=1, column=1, sticky=tk.W, pady=(5, 5), padx=(5, 0))

        self.delay_var = tk.IntVar(value=2)
        self.delay_spinbox = ttk.Spinbox(delay_frame, from_=1, to=30,
                                        textvariable=self.delay_var, width=5)
        self.delay_spinbox.pack(side=tk.LEFT)
        ttk.Label(delay_frame, text="seconds", foreground="gray").pack(side=tk.LEFT, padx=(5, 0))

        # --- Row 2: Hint label spanning both columns ---
        ttk.Label(filename_frame,
                text="The .bat extension is added automatically. Delay applies between each program launch.",
                font=('Arial', 9), foreground="gray").grid(row=2, column=0, columnspan=2, sticky=tk.W)

        filename_frame.columnconfigure(1, weight=1)
            
        # Summary frame
        summary_frame = ttk.LabelFrame(build_frame, text="Summary", padding="10")
        summary_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.summary_text = scrolledtext.ScrolledText(summary_frame, wrap=tk.WORD, height=10)
        self.summary_text.pack(fill=tk.BOTH, expand=True)
        
        # Update summary
        self.update_summary()
        
        # Build button
        button_frame = ttk.Frame(build_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        #Everything updates automatically now
        self.build_button = ttk.Button(button_frame, text="⚙️Build Batch File⚙️", 
                                    command=self.build_batch, style="Accent.TButton")
        self.build_button.pack()
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(build_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
        
        # Status label
        self.build_status_label = ttk.Label(build_frame, text="")
        self.build_status_label.pack()
        

    def update_batch_filename(self):
        """Update the batch file name from user input."""
        new_name = self.batch_filename_var.get().strip()
        
        if not new_name:
            messagebox.showwarning("Warning", "Please enter a file name")
            return
        
        # Remove any .bat extension if user included it
        if new_name.lower().endswith('.bat'):
            new_name = new_name[:-4]
        
        # Validate file name
        if not re.match(r'^[a-zA-Z0-9_\- ]+$', new_name):
            messagebox.showerror("Error", 
                                "Invalid file name. Use only letters, numbers, spaces, hyphens, and underscores.")
            # Reset to previous valid name
            self.batch_filename_var.set("Open_Stream")
            return
        
        # Update the variable
        self.batch_filename_var.set(new_name)
        self.update_status(f"Batch file name updated to: {new_name}.bat")

    def create_status_bar(self):
        self.status_bar = ttk.Label(self.main_frame, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

    # Helper methods
    def update_status(self, message):
        self.status_bar.config(text=message)
        self.root.update_idletasks()

    def update_obs_status(self):
        if self.obs_path:
            self.obs_status_label.config(text=f"✓ {self.obs_path}", style='Success.TLabel')
            # Show the clear button when OBS is configured
            self.clear_obs_button.pack(side=tk.RIGHT)
        else:
            self.obs_status_label.config(text="No OBS configured", style='Error.TLabel')
            # Hide the clear button when no OBS is configured
            self.clear_obs_button.pack_forget()


    def refresh_view(self):
        # Enable the text widget for editing
        self.view_text.config(state='normal')
        self.view_text.delete(1.0, tk.END)
        
        # Add OBS
        self.view_text.insert(tk.END, "📷 BROADCAST SOFTWARE:\n")
        if self.obs_path:
            self.view_text.insert(tk.END, f"  • OBS: {self.obs_path}\n")
        else:
            self.view_text.insert(tk.END, "  • No OBS configured\n")
        self.view_text.insert(tk.END, "\n")
        
        # Add Steam games
        self.view_text.insert(tk.END, "🎮 STEAM GAMES/PROGRAMS:\n")
        if self.steam_dictionary:
            for name, url in self.steam_dictionary.items():
                self.view_text.insert(tk.END, f"  • {name}: {url}\n")
        else:
            self.view_text.insert(tk.END, "  • No Steam games/programs added\n")
        self.view_text.insert(tk.END, "\n")
        
        # Add Programs
        self.view_text.insert(tk.END, "💻 OTHER PROGRAMS:\n")
        if self.program_shortcut_dictionary:
            for name, path in self.program_shortcut_dictionary.items():
                self.view_text.insert(tk.END, f"  • {name}: {path}\n")
        else:
            self.view_text.insert(tk.END, "  • No programs added\n")
        self.view_text.insert(tk.END, "\n")
        
        # Add MS Store apps
        self.view_text.insert(tk.END, "🏪 MICROSOFT STORE APPS:\n")
        if self.ms_store_dictionary:
            for name, aumid in self.ms_store_dictionary.items():
                self.view_text.insert(tk.END, f"  • {name}: {aumid}\n")
        else:
            self.view_text.insert(tk.END, "  • No MS Store apps added\n")
        self.view_text.insert(tk.END, "\n")

        # Add Web pages
        self.view_text.insert(tk.END, "🌐 WEB PAGES:\n")
        if self.web_dictionary:
            for name, url in self.web_dictionary.items():
                self.view_text.insert(tk.END, f"  • {name}: {url}\n")
        else:
            self.view_text.insert(tk.END, "  • No web pages added\n")
        
        # Disable the text widget to prevent editing
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

        self.summary_text.insert(tk.END, f"LAUNCH LIST\n")
        self.summary_text.insert(tk.END, f"{'='*40}\n\n")
        self.summary_text.insert(tk.END, f"Total items in launch list: {total_items}\n\n")

        self.summary_text.insert(tk.END, f"OBS: {'✓ Configured' if self.obs_path else 'X Not configured'}\n")
        self.summary_text.insert(tk.END, f"Steam Games/Programs: {len(self.steam_dictionary)}\n")
        self.summary_text.insert(tk.END, f"Programs: {len(self.program_shortcut_dictionary)}\n")
        self.summary_text.insert(tk.END, f"MS Store Apps: {len(self.ms_store_dictionary)}\n")
        self.summary_text.insert(tk.END, f"Web Pages: {len(self.web_dictionary)}\n\n")
        
        if total_items > 0:
            self.summary_text.insert(tk.END, "Ready to build batch file!\n")
        else:
            self.summary_text.insert(tk.END, "! Please add at least one program/game/page before building the batch file !\n")

        self.summary_text.config(state='disabled')


    # NEW: Auto-refresh method
    def auto_refresh(self):
        """Automatically refresh the view list and build summary."""
        self.refresh_view()
        self.update_summary()

    # NEW: Clean path by removing quotation marks
    def clean_path(self, path):
        """Remove quotation marks and strip whitespace from a path."""
        if not path:
            return path
        # Remove both single and double quotes
        cleaned = path.replace('"', '').replace("'", "")
        # Strip whitespace
        cleaned = cleaned.strip()
        return cleaned

    # File browsing methods
    def browse_obs(self):
        filename = filedialog.askopenfilename(
            title="Select OBS Executable",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
        )
        if filename:
            # Clean the path
            cleaned_path = self.clean_path(filename)
            path = Path(cleaned_path)
            if path.exists() and path.suffix.lower() == '.exe':
                self.obs_path_var.set(str(path))
                self.obs_path = str(path)
                self.update_obs_status()
                self.update_status("OBS path updated successfully")
                # Auto-refresh after adding
                self.auto_refresh()
            else:
                messagebox.showerror("Error", "Please select a valid .exe file")

    # NEW: Set OBS path from text entry
    def set_obs_path(self):
        path_input = self.obs_path_var.get().strip()
        
        if not path_input:
            messagebox.showwarning("Warning", "Please enter a path for OBS")
            return
        
        # Clean the path (remove quotes)
        cleaned_path = self.clean_path(path_input)
        
        try:
            path = Path(cleaned_path)
            
            # Check if path exists
            if not path.exists():
                messagebox.showerror("Error", f"The path '{cleaned_path}' does not exist!")
                return
            
            # Check if it's a file
            if not path.is_file():
                messagebox.showerror("Error", f"'{cleaned_path}' is not a file!")
                return
            
            # Check if it's an .exe file
            if path.suffix.lower() != '.exe':
                messagebox.showerror("Error", f"'{cleaned_path}' is not an .exe file!")
                return
            
            # Set the OBS path
            self.obs_path = str(path)
            self.update_obs_status()
            self.update_status("OBS set successfully from manual patch input")
            # Auto-refresh after adding
            self.auto_refresh()
            
        except Exception as e:
            messagebox.showerror("Error", f"Invalid path: {str(e)}")

    def browse_program(self):
        filename = filedialog.askopenfilename(
            title="Select Program Executable",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
        )
        if filename:
            # Clean the path
            cleaned_path = self.clean_path(filename)
            path = Path(cleaned_path)
            if path.exists() and path.suffix.lower() == '.exe':
                self.prog_path_var.set(str(path))
            else:
                messagebox.showerror("Error", "Please select a valid .exe file")

    # Add methods (updated to clean paths and auto-refresh)
    def add_program(self):
        name = self.prog_name_var.get().strip()
        path_input = self.prog_path_var.get().strip()
        
        if not name:
            messagebox.showwarning("Warning", "Please enter a program name")
            return
        
        if not path_input:
            messagebox.showwarning("Warning", "Please enter an executable path")
            return
        
        # Clean the path (remove quotes)
        cleaned_path = self.clean_path(path_input)
        
        try:
            path_obj = Path(cleaned_path)
            
            if not path_obj.exists():
                messagebox.showerror("Error", f"The path '{cleaned_path}' does not exist!")
                return
            
            if not path_obj.is_file():
                messagebox.showerror("Error", f"'{cleaned_path}' is not a file!")
                return
            
            if path_obj.suffix.lower() != '.exe':
                messagebox.showerror("Error", f"'{cleaned_path}' is not an .exe file!")
                return
            
            # Store the cleaned path
            self.program_shortcut_dictionary[name] = str(path_obj)
            
            # Update treeview
            self.programs_tree.insert('', tk.END, values=(name, str(path_obj)))
            
            # Clear entries
            self.prog_name_var.set('')
            self.prog_path_var.set('')
            
            self.update_status(f"Added program: {name}")
            # Auto-refresh after adding
            self.auto_refresh()
            
        except Exception as e:
            messagebox.showerror("Error", f"Invalid path: {str(e)}")

    def add_steam(self):
        name = self.steam_name_var.get().strip()
        url = self.steam_url_var.get().strip()
        
        if not name:
            messagebox.showwarning("Warning", "Please enter a Steam game/program name")
            return
        
        if not url:
            messagebox.showwarning("Warning", "Please enter a Steam URL")
            return
        
        # Clean the URL (remove quotes)
        cleaned_url = self.clean_path(url)
        
        is_valid, message = validate_steam_url(cleaned_url)
        
        if not is_valid:
            messagebox.showerror("Error", f"Invalid Steam URL: {message}")
            return
        
        self.steam_dictionary[name] = cleaned_url
        
        # Update treeview
        self.steam_tree.insert('', tk.END, values=(name, cleaned_url))
        
        # Clear entries
        self.steam_name_var.set('')
        self.steam_url_var.set('')
        
        self.update_status(f"Added Steam game/program: {name}")
        # Auto-refresh after adding
        self.auto_refresh()

    def search_ms_store_apps(self):
        query = self.ms_search_var.get().strip()
        if not query:
            messagebox.showwarning("Warning", "Please enter an app name to search for")
            return

        self.ms_search_button.config(state='disabled')
        self.ms_search_status.config(text="Searching...", foreground="gray")
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
            # ConvertTo-Json returns a dict for a single result, list for multiple
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
            self.ms_search_status.config(text="No apps found. Try a different name.", foreground="red")
            return
        self.ms_search_status.config(text=f"{len(apps)} result(s) found. Select one and click 'Add Selected App'.", foreground="green")
        for name, aumid in apps:
            self.ms_results_tree.insert('', tk.END, values=(name, aumid))

    def _ms_search_error(self, error):
        self.ms_search_button.config(state='normal')
        self.ms_search_status.config(text=f"Search failed: {error}", foreground="red")

    def add_ms_store_from_results(self):
        selected = self.ms_results_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an app from the search results")
            return

        values = self.ms_results_tree.item(selected[0], 'values')
        name, aumid = values[0], values[1]

        if name in self.ms_store_dictionary:
            messagebox.showinfo("Info", f"'{name}' is already in the launch list.")
            return

        self.ms_store_dictionary[name] = aumid
        self.ms_store_tree.insert('', tk.END, values=(name, aumid))
        self.update_status(f"Added MS Store app: {name}")
        self.auto_refresh()

    def add_web(self):
        name = self.web_name_var.get().strip()
        url = self.web_url_var.get().strip()
        
        if not name:
            messagebox.showwarning("Warning", "Please enter a page name")
            return
        
        if not url:
            messagebox.showwarning("Warning", "Please enter a web URL")
            return
        
        # Clean the URL (remove quotes)
        cleaned_url = self.clean_path(url)
        
        if not cleaned_url.startswith(('http://', 'https://')):
            cleaned_url = 'https://' + cleaned_url
        
        self.web_dictionary[name] = cleaned_url
        
        # Update treeview
        self.web_tree.insert('', tk.END, values=(name, cleaned_url))
        
        # Clear entries
        self.web_name_var.set('')
        self.web_url_var.set('')
        
        self.update_status(f"Added web page: {name}")
        # Auto-refresh after adding
        self.auto_refresh()

    # Remove methods (updated to auto-refresh)
    def remove_program(self):
        selected = self.programs_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a program to remove")
            return
        
        for item in selected:
            values = self.programs_tree.item(item, 'values')
            name = values[0]
            if name in self.program_shortcut_dictionary:
                del self.program_shortcut_dictionary[name]
            self.programs_tree.delete(item)
        
        self.update_status("Program removed")
        # Auto-refresh after removing
        self.auto_refresh()

    def remove_steam(self):
        selected = self.steam_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a Steam game/program to remove")
            return
        
        for item in selected:
            values = self.steam_tree.item(item, 'values')
            name = values[0]
            if name in self.steam_dictionary:
                del self.steam_dictionary[name]
            self.steam_tree.delete(item)
        
        self.update_status("Steam game/program removed")
        # Auto-refresh after removing
        self.auto_refresh()

    def remove_ms_store(self):
        selected = self.ms_store_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an app to remove")
            return

        for item in selected:
            values = self.ms_store_tree.item(item, 'values')
            name = values[0]
            if name in self.ms_store_dictionary:
                del self.ms_store_dictionary[name]
            self.ms_store_tree.delete(item)

        self.update_status("MS Store app removed")
        self.auto_refresh()

    def remove_web(self):
        selected = self.web_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a web page to remove")
            return
        
        for item in selected:
            values = self.web_tree.item(item, 'values')
            name = values[0]
            if name in self.web_dictionary:
                del self.web_dictionary[name]
            self.web_tree.delete(item)
        
        self.update_status("Web page removed")
        # Auto-refresh after removing
        self.auto_refresh()

    # Build method
    def build_batch(self):
        # Check if there's anything to build
        if not self.obs_path and not self.steam_dictionary and not self.program_shortcut_dictionary and not self.web_dictionary and not self.ms_store_dictionary:
            messagebox.showwarning("Warning", "Please add at least one item(program/game/page) before building")
            return
        
        # Confirm with user
        response = messagebox.askyesno("Confirm Build", 
                                      "This will create a batch(.bat) file that launches all the added items.\n\n"
                                      "Do you wish to proceed?")
        if not response:
            return
        
        # Disable build button during process
        self.build_button.config(state='disabled')
        self.progress_var.set(0)
        self.build_status_label.config(text="Building .bat file...")
        
        # Run in separate thread to keep GUI responsive
        thread = threading.Thread(target=self._build_batch_thread)
        thread.daemon = True
        thread.start()

    def _build_batch_thread(self):
        # Initialize COM for this thread
        import pythoncom
        pythoncom.CoInitialize()
        
        try:
            # Step 1: Create shortcut folder
            self.root.after(0, lambda: self.update_status("Creating shortcut folder..."))
            self.root.after(0, lambda: self.progress_var.set(20))
            
            shortcut_folder_dir = make_empty_shortcut_folder()
            
            # Step 2: Fill shortcut folder and get OBS shortcut path
            self.root.after(0, lambda: self.update_status("Creating shortcuts..."))
            self.root.after(0, lambda: self.progress_var.set(40))
            
            obs_shortcut_path = fill_shortcut_folder(
                shortcut_dir=shortcut_folder_dir,
                S_dictionary=self.program_shortcut_dictionary if self.program_shortcut_dictionary else None,
                steam_dict=self.steam_dictionary if self.steam_dictionary else None,
                obs_dir=self.obs_path,
                web_dict=self.web_dictionary if self.web_dictionary else None,
                ms_store_dict=self.ms_store_dictionary if self.ms_store_dictionary else None
            )
            
            # Step 3: Create batch file with OBS shortcut path
            self.root.after(0, lambda: self.update_status("Creating batch file..."))
            self.root.after(0, lambda: self.progress_var.set(60))

            batch_filename = self.batch_filename_var.get() + ".bat"
            delay_seconds = self.delay_var.get()  

            create_bat_file(
                self.program_shortcut_dictionary,
                shortcut_folder_dir,
                obs_short_dir=obs_shortcut_path,  # Pass the actual OBS shortcut path
                batch_filename=batch_filename,  #Name Parameter
                delay_seconds=delay_seconds #Selected delay paameter
            )
            
            # Step 4: Complete
            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(0, lambda: self.update_status("Batch file created successfully!"))
            self.root.after(0, lambda: self.build_status_label.config(text="✓ Batch file created successfully!"))
            
            # Show success message
            self.root.after(0, lambda: messagebox.showinfo("Success", 
                f"Batch file created successfully!\n\n"
                f"The file '{batch_filename}' has been created in the current directory.\n"
                f"Double-click it to launch all your programs!\n\n"
                f"IMPORTANT: Do not move the 'Shortcuts_for_bat' folder, or the batch file won't work."))
            
        except Exception as e:
            # Capture the exception in the local scope of the lambda
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self.update_status(f"Error: {msg}"))
            self.root.after(0, lambda msg=error_msg: self.build_status_label.config(text=f"X Error: {msg}", style='Error.TLabel'))
            self.root.after(0, lambda msg=error_msg: messagebox.showerror("Error", f"An error occurred:\n{msg}"))
        finally:
            # Uninitialize COM
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self.build_button.config(state='normal'))

# Keep your original functions (they should remain unchanged)
def make_empty_shortcut_folder():
    current_folder = os.getcwd()
    shortcut_folder_path = f"{current_folder}\\Shortcuts_for_bat"
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

def create_bat_file(S_dictionary, shortcut_dir, obs_short_dir=None, batch_filename="Open_Stream.bat", delay_seconds = 2):
    # And use batch_filename directly:
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
                return True, "Valid Steam URL."
        
        return False, "Invalid Steam URL format. Should be like: steam://rungameid/123456."
    
    else:
        return False, "Steam URL must start with 'steam://'."
    
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
        print(f"Could not load icon: {e}")
    
    app = StreamBatGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()