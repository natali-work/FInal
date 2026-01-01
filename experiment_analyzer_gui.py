"""
Experiment Analyzer GUI
A graphical interface for processing and analyzing respiratory experiment data.
Supports single experiment analysis and multi-experiment comparison.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import threading
from datetime import datetime
import traceback

# Import processing functions
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from scipy import stats
import gc
import warnings
import re
import glob
from openpyxl import load_workbook
warnings.filterwarnings('ignore')


class ExperimentAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Experiment Analyzer - Respiratory Data Analysis")
        self.root.geometry("1000x850")
        self.root.minsize(900, 750)
        
        # Configure color scheme
        self.colors = {
            'bg': '#F5F5F5',
            'primary': '#2C3E50',
            'secondary': '#3498DB',
            'accent': '#E74C3C',
            'success': '#27AE60',
            'warning': '#F39C12',
            'frame_bg': '#FFFFFF',
            'text': '#2C3E50',
            'text_light': '#7F8C8D',
        }
        
        # Set window background
        self.root.configure(bg=self.colors['bg'])
        
        # Variables for file paths (single experiment mode)
        self.baseline_file = tk.StringVar()
        self.experiment_file = tk.StringVar()
        self.output_dir = tk.StringVar(value=os.getcwd())
        
        # Multi-experiment mode
        self.experiments_list = []  # List of {'name': str, 'baseline': str, 'experiment': str}
        
        # Processing parameters
        self.pre_treatment_ignore_minutes = tk.DoubleVar(value=1.5)
        self.post_treatment_ignore_minutes = tk.DoubleVar(value=2.0)
        self.pre_treatment_avg_seconds = tk.IntVar(value=60)
        self.post_treatment_avg_seconds = tk.IntVar(value=20)
        
        # Mode variable
        self.mode = tk.StringVar(value="single")
        
        # Configure styles
        self.setup_styles()
        
        # Build the GUI
        self.create_widgets()
        
    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling"""
        if sys.platform == "win32":
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")
        
    def _bind_mousewheel(self, event):
        """Bind mouse wheel to canvas"""
        if sys.platform == "win32":
            self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        else:
            self.canvas.bind_all("<Button-4>", self._on_mousewheel)
            self.canvas.bind_all("<Button-5>", self._on_mousewheel)
        
    def _unbind_mousewheel(self, event):
        """Unbind mouse wheel from canvas"""
        if sys.platform == "win32":
            self.canvas.unbind_all("<MouseWheel>")
        else:
            self.canvas.unbind_all("<Button-4>")
            self.canvas.unbind_all("<Button-5>")
        
    def setup_styles(self):
        """Configure ttk styles for a modern look"""
        style = ttk.Style()
        
        # Try to use a modern theme
        try:
            style.theme_use('vista')
        except:
            try:
                style.theme_use('clam')
            except:
                pass
        
        # Configure frame styles
        style.configure('Title.TLabel', 
                       font=('Segoe UI', 18, 'bold'),
                       background=self.colors['bg'],
                       foreground=self.colors['primary'])
        
        style.configure('Subtitle.TLabel',
                       font=('Segoe UI', 9),
                       background=self.colors['bg'],
                       foreground=self.colors['text_light'])
        
        style.configure('Header.TLabel',
                       font=('Segoe UI', 10, 'bold'),
                       foreground=self.colors['primary'])
        
        # Configure button styles
        style.configure('Primary.TButton',
                       font=('Segoe UI', 10, 'bold'),
                       padding=10)
        
        style.map('Primary.TButton',
                 background=[('active', '#1A1A1A'),
                           ('!active', '#000000')],
                 foreground=[('active', 'black'),
                            ('!active', 'black')])
        
        style.configure('Secondary.TButton',
                       font=('Segoe UI', 9),
                       padding=8)
        
        # Configure frame styles
        style.configure('Card.TLabelframe',
                       background=self.colors['frame_bg'],
                       borderwidth=1,
                       relief='solid')
        
        style.configure('Card.TLabelframe.Label',
                       font=('Segoe UI', 10, 'bold'),
                       foreground=self.colors['primary'],
                       background=self.colors['frame_bg'])
        
        # Configure entry styles
        style.configure('Modern.TEntry',
                       fieldbackground='white',
                       borderwidth=1,
                       relief='solid',
                       padding=5)
        
        # Configure status bar
        style.configure('Status.TLabel',
                       font=('Segoe UI', 9),
                       background=self.colors['primary'],
                       foreground='white',
                       padding=5)
        
    def create_widgets(self):
        # Create scrollable frame
        # Create canvas and scrollbar container
        canvas_container = tk.Frame(self.root, bg=self.colors['bg'])
        canvas_container.pack(fill=tk.BOTH, expand=True)
        
        self.canvas = tk.Canvas(canvas_container, bg=self.colors['bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=self.colors['bg'])
        
        # Function to update scroll region
        def update_scroll_region(event=None):
            self.canvas.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        self.scrollable_frame.bind("<Configure>", lambda e: update_scroll_region())
        
        canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        # Update canvas width when window is resized
        def configure_canvas_width(event):
            canvas_width = event.width
            self.canvas.itemconfig(canvas_window, width=canvas_width)
        self.canvas.bind('<Configure>', configure_canvas_width)
        
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas and scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mouse wheel
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)
        
        # Main container with padding (inside scrollable frame)
        main_frame = tk.Frame(self.scrollable_frame, bg=self.colors['bg'], padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header section with title
        header_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        title_label = ttk.Label(header_frame, text="🔬 Experiment Analyzer", 
                                style='Title.TLabel')
        title_label.pack(anchor=tk.W)
        
        subtitle_label = ttk.Label(header_frame, 
                                   text="Respiratory Data Processing & Analysis Tool",
                                   style='Subtitle.TLabel')
        subtitle_label.pack(anchor=tk.W, pady=(2, 0))
        
        # ============ Mode Selection ============
        mode_frame = ttk.LabelFrame(main_frame, text="📊 Analysis Mode", 
                                   style='Card.TLabelframe', padding="12")
        mode_frame.pack(fill=tk.X, pady=(0, 12))
        
        mode_inner = tk.Frame(mode_frame, bg=self.colors['frame_bg'])
        mode_inner.pack(fill=tk.X)
        
        ttk.Radiobutton(mode_inner, text="🔬 Single Experiment", variable=self.mode, 
                        value="single", command=self.toggle_mode).pack(side=tk.LEFT, padx=15, pady=5)
        ttk.Radiobutton(mode_inner, text="📈 Compare Multiple Experiments", variable=self.mode,
                        value="multi", command=self.toggle_mode).pack(side=tk.LEFT, padx=15, pady=5)
        
        # ============ Single Experiment Frame ============
        self.single_frame = ttk.LabelFrame(main_frame, text="📁 Single Experiment Files", 
                                           style='Card.TLabelframe', padding="15")
        # Will be packed after output_frame is created
        
        # Baseline file
        ttk.Label(self.single_frame, text="Baseline File:", style='Header.TLabel').grid(
            row=0, column=0, sticky=tk.W, pady=8, padx=(0, 10))
        baseline_entry = ttk.Entry(self.single_frame, textvariable=self.baseline_file, 
                                   width=55, style='Modern.TEntry')
        baseline_entry.grid(row=0, column=1, padx=5, pady=8, sticky=tk.EW)
        ttk.Button(self.single_frame, text="📂 Browse", command=self.browse_baseline,
                   style='Secondary.TButton').grid(row=0, column=2, padx=(5, 0))
        
        # Experiment file
        ttk.Label(self.single_frame, text="Experiment File:", style='Header.TLabel').grid(
            row=1, column=0, sticky=tk.W, pady=8, padx=(0, 10))
        experiment_entry = ttk.Entry(self.single_frame, textvariable=self.experiment_file, 
                                    width=55, style='Modern.TEntry')
        experiment_entry.grid(row=1, column=1, padx=5, pady=8, sticky=tk.EW)
        ttk.Button(self.single_frame, text="📂 Browse", command=self.browse_experiment,
                   style='Secondary.TButton').grid(row=1, column=2, padx=(5, 0))
        
        self.single_frame.columnconfigure(1, weight=1)
        
        # ============ Multi Experiment Frame ============
        self.multi_frame = ttk.LabelFrame(main_frame, text="📊 Multiple Experiments", 
                                          style='Card.TLabelframe', padding="15")
        # Don't pack yet - will be shown when mode changes
        
        # File upload section - make it very visible and simple
        upload_label = tk.Label(self.multi_frame, 
                               text="📁 Add Experiment Files:", 
                               font=('Segoe UI', 11, 'bold'),
                               fg=self.colors['primary'],
                               bg=self.colors['frame_bg'],
                               anchor=tk.W)
        upload_label.pack(fill=tk.X, pady=(0, 10))
        
        # Experiment name
        name_frame = tk.Frame(self.multi_frame, bg=self.colors['frame_bg'])
        name_frame.pack(fill=tk.X, pady=5)
        ttk.Label(name_frame, text="Experiment Name:", style='Header.TLabel').pack(side=tk.LEFT, padx=(0, 10))
        self.quick_name_var = tk.StringVar()
        name_entry = ttk.Entry(name_frame, textvariable=self.quick_name_var, width=40, 
                              style='Modern.TEntry')
        name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Baseline file
        baseline_frame = tk.Frame(self.multi_frame, bg=self.colors['frame_bg'])
        baseline_frame.pack(fill=tk.X, pady=5)
        ttk.Label(baseline_frame, text="Baseline File:", style='Header.TLabel').pack(side=tk.LEFT, padx=(0, 10))
        self.quick_baseline_var = tk.StringVar()
        baseline_entry = ttk.Entry(baseline_frame, textvariable=self.quick_baseline_var, width=40, 
                                   style='Modern.TEntry')
        baseline_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(baseline_frame, text="📂 Browse", 
                   command=lambda: self.browse_quick_file(self.quick_baseline_var),
                   style='Secondary.TButton').pack(side=tk.LEFT, padx=5)
        
        # Experiment file
        experiment_frame = tk.Frame(self.multi_frame, bg=self.colors['frame_bg'])
        experiment_frame.pack(fill=tk.X, pady=5)
        ttk.Label(experiment_frame, text="Experiment File:", style='Header.TLabel').pack(side=tk.LEFT, padx=(0, 10))
        self.quick_experiment_var = tk.StringVar()
        experiment_entry = ttk.Entry(experiment_frame, textvariable=self.quick_experiment_var, width=40, 
                                    style='Modern.TEntry')
        experiment_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(experiment_frame, text="📂 Browse", 
                   command=lambda: self.browse_quick_file(self.quick_experiment_var),
                   style='Secondary.TButton').pack(side=tk.LEFT, padx=5)
        
        # Add button
        add_btn_frame = tk.Frame(self.multi_frame, bg=self.colors['frame_bg'])
        add_btn_frame.pack(fill=tk.X, pady=(10, 15))
        ttk.Button(add_btn_frame, text="➕ Add Experiment to List", 
                   command=self.add_quick_experiment,
                   style='Secondary.TButton').pack(side=tk.LEFT)
        
        # Separator
        separator = ttk.Separator(self.multi_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=(0, 10))
        
        # Experiments list label
        list_label = tk.Label(self.multi_frame, text="📋 Experiments List:", 
                             font=('Segoe UI', 10, 'bold'),
                             fg=self.colors['primary'],
                             bg=self.colors['frame_bg'],
                             anchor=tk.W)
        list_label.pack(fill=tk.X, pady=(0, 5))
        
        list_frame = ttk.Frame(self.multi_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Treeview for experiments
        columns = ('name', 'baseline', 'experiment')
        self.exp_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=5)
        self.exp_tree.heading('name', text='Experiment Name')
        self.exp_tree.heading('baseline', text='Baseline File')
        self.exp_tree.heading('experiment', text='Experiment File')
        self.exp_tree.column('name', width=150)
        self.exp_tree.column('baseline', width=250)
        self.exp_tree.column('experiment', width=250)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.exp_tree.yview)
        self.exp_tree.configure(yscrollcommand=scrollbar.set)
        
        self.exp_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Buttons for managing experiments
        btn_frame = tk.Frame(self.multi_frame, bg=self.colors['frame_bg'])
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(btn_frame, text="➕ Add Experiment", command=self.add_experiment,
                   style='Secondary.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="➖ Remove Selected", command=self.remove_experiment,
                   style='Secondary.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ Clear All", command=self.clear_experiments,
                   style='Secondary.TButton').pack(side=tk.LEFT, padx=5)
        
        # ============ Output Directory ============
        self.output_frame = ttk.LabelFrame(main_frame, text="💾 Output Directory", 
                                     style='Card.TLabelframe', padding="15")
        self.output_frame.pack(fill=tk.X, pady=(0, 12))
        
        ttk.Label(self.output_frame, text="Output Directory:", style='Header.TLabel').grid(
            row=0, column=0, sticky=tk.W, pady=8, padx=(0, 10))
        output_entry = ttk.Entry(self.output_frame, textvariable=self.output_dir, 
                                 width=55, style='Modern.TEntry')
        output_entry.grid(row=0, column=1, padx=5, pady=8, sticky=tk.EW)
        ttk.Button(self.output_frame, text="📂 Browse", command=self.browse_output,
                   style='Secondary.TButton').grid(row=0, column=2, padx=(5, 0))
        self.output_frame.columnconfigure(1, weight=1)
        
        # Pack single_frame initially (default mode is "single")
        self.single_frame.pack(fill=tk.X, pady=(0, 12), before=self.output_frame)
        
        # ============ Parameters Frame ============
        params_frame = ttk.LabelFrame(main_frame, text="⚙️ Processing Parameters", 
                                     style='Card.TLabelframe', padding="15")
        params_frame.pack(fill=tk.X, pady=(0, 12))
        
        # Pre-treatment settings
        pre_frame = ttk.LabelFrame(params_frame, text="Pre-Treatment Settings", 
                                  style='Card.TLabelframe', padding="10")
        pre_frame.pack(fill=tk.X, pady=(0, 8))
        
        ttk.Label(pre_frame, text="Minutes to ignore from start:", 
                  style='Header.TLabel').grid(row=0, column=0, sticky=tk.W, padx=8, pady=5)
        pre_ignore_spin = ttk.Spinbox(pre_frame, from_=0, to=10, increment=0.5, 
                                       textvariable=self.pre_treatment_ignore_minutes, 
                                       width=12, style='Modern.TEntry')
        pre_ignore_spin.grid(row=0, column=1, padx=8, pady=5)
        ttk.Label(pre_frame, text="minutes", 
                 foreground=self.colors['text_light']).grid(row=0, column=2, sticky=tk.W, padx=(0, 8))
        
        ttk.Label(pre_frame, text="Averaging interval:", 
                 style='Header.TLabel').grid(row=1, column=0, sticky=tk.W, padx=8, pady=5)
        pre_avg_spin = ttk.Spinbox(pre_frame, from_=10, to=120, increment=10,
                                    textvariable=self.pre_treatment_avg_seconds, 
                                    width=12, style='Modern.TEntry')
        pre_avg_spin.grid(row=1, column=1, padx=8, pady=5)
        ttk.Label(pre_frame, text="seconds", 
                 foreground=self.colors['text_light']).grid(row=1, column=2, sticky=tk.W, padx=(0, 8))
        
        # Post-treatment settings
        post_frame = ttk.LabelFrame(params_frame, text="Post-Treatment Settings", 
                                   style='Card.TLabelframe', padding="10")
        post_frame.pack(fill=tk.X, pady=(0, 0))
        
        ttk.Label(post_frame, text="Minutes to ignore from start:", 
                 style='Header.TLabel').grid(row=0, column=0, sticky=tk.W, padx=8, pady=5)
        post_ignore_spin = ttk.Spinbox(post_frame, from_=0, to=10, increment=0.5,
                                        textvariable=self.post_treatment_ignore_minutes, 
                                        width=12, style='Modern.TEntry')
        post_ignore_spin.grid(row=0, column=1, padx=8, pady=5)
        ttk.Label(post_frame, text="minutes", 
                 foreground=self.colors['text_light']).grid(row=0, column=2, sticky=tk.W, padx=(0, 8))
        
        ttk.Label(post_frame, text="Averaging interval:", 
                 style='Header.TLabel').grid(row=1, column=0, sticky=tk.W, padx=8, pady=5)
        post_avg_spin = ttk.Spinbox(post_frame, from_=10, to=120, increment=10,
                                     textvariable=self.post_treatment_avg_seconds, 
                                     width=12, style='Modern.TEntry')
        post_avg_spin.grid(row=1, column=1, padx=8, pady=5)
        ttk.Label(post_frame, text="seconds", 
                 foreground=self.colors['text_light']).grid(row=1, column=2, sticky=tk.W, padx=(0, 8))
        
        # ============ Action Buttons ============
        button_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.run_button = ttk.Button(button_frame, text="▶️ Run Analysis", 
                                      command=self.run_analysis, style='Primary.TButton')
        self.run_button.pack(side=tk.LEFT, padx=(0, 8))
        
        ttk.Button(button_frame, text="🗑️ Clear Log", command=self.clear_log,
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=4)
        ttk.Button(button_frame, text="📂 Open Output Folder", command=self.open_output_folder,
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=4)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate', length=400)
        self.progress.pack(fill=tk.X, pady=(0, 10))
        
        # ============ Log Output ============
        log_frame = ttk.LabelFrame(main_frame, text="📋 Log Output", 
                                  style='Card.TLabelframe', padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, wrap=tk.WORD,
                                                   font=('Consolas', 9),
                                                   bg='#FAFAFA',
                                                   fg=self.colors['text'],
                                                   insertbackground=self.colors['primary'],
                                                   selectbackground=self.colors['secondary'],
                                                   selectforeground='white',
                                                   relief='flat',
                                                   borderwidth=1)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Status bar (outside scrollable area, at bottom of window)
        status_frame = tk.Frame(self.root, bg=self.colors['primary'], height=30)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        status_frame.pack_propagate(False)
        
        self.status_var = tk.StringVar(value="✓ Ready")
        status_bar = tk.Label(status_frame, textvariable=self.status_var, 
                             bg=self.colors['primary'],
                             fg='white',
                             font=('Segoe UI', 9),
                             anchor=tk.W,
                             padx=10)
        status_bar.pack(fill=tk.BOTH, expand=True)
        
    def toggle_mode(self):
        """Toggle between single and multi-experiment modes"""
        if self.mode.get() == "single":
            self.multi_frame.pack_forget()
            self.single_frame.pack(fill=tk.X, pady=(0, 12), before=self.output_frame)
        else:
            self.single_frame.pack_forget()
            self.multi_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 12), before=self.output_frame)
            
    def add_experiment(self):
        """Add a new experiment to the list"""
        dialog = AddExperimentDialog(self.root)
        self.root.wait_window(dialog.top)
        
        if dialog.result:
            self.experiments_list.append(dialog.result)
            self.exp_tree.insert('', 'end', values=(
                dialog.result['name'],
                os.path.basename(dialog.result['baseline']),
                os.path.basename(dialog.result['experiment'])
            ))
            
    def remove_experiment(self):
        """Remove selected experiment from the list"""
        selected = self.exp_tree.selection()
        if selected:
            idx = self.exp_tree.index(selected[0])
            self.exp_tree.delete(selected[0])
            del self.experiments_list[idx]
            
    def clear_experiments(self):
        """Clear all experiments from the list"""
        for item in self.exp_tree.get_children():
            self.exp_tree.delete(item)
        self.experiments_list.clear()
    
    def browse_quick_file(self, var):
        """Browse for a file and set it to the given variable"""
        filename = filedialog.askopenfilename(
            title="Select File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            var.set(filename)
            # Auto-fill experiment name if it's a baseline file and name is empty
            if var == self.quick_baseline_var and not self.quick_name_var.get():
                name = os.path.splitext(os.path.basename(filename))[0]
                self.quick_name_var.set(name.replace('_', ' ').replace(' baseline', ''))
    
    def add_quick_experiment(self):
        """Add experiment from quick add fields"""
        name = self.quick_name_var.get().strip()
        baseline = self.quick_baseline_var.get().strip()
        experiment = self.quick_experiment_var.get().strip()
        
        if not name:
            messagebox.showerror("Error", "Please enter an experiment name")
            return
        if not baseline:
            messagebox.showerror("Error", "Please select a baseline file")
            return
        if not experiment:
            messagebox.showerror("Error", "Please select an experiment file")
            return
        if not os.path.exists(baseline):
            messagebox.showerror("Error", f"Baseline file not found: {baseline}")
            return
        if not os.path.exists(experiment):
            messagebox.showerror("Error", f"Experiment file not found: {experiment}")
            return
        
        # Add to list
        exp_data = {
            'name': name,
            'baseline': baseline,
            'experiment': experiment,
        }
        self.experiments_list.append(exp_data)
        self.exp_tree.insert('', 'end', values=(
            name,
            os.path.basename(baseline),
            os.path.basename(experiment)
        ))
        
        # Clear the quick add fields
        self.quick_name_var.set('')
        self.quick_baseline_var.set('')
        self.quick_experiment_var.set('')
        
        messagebox.showinfo("Success", f"Experiment '{name}' added to list!")
        
    def browse_baseline(self):
        filename = filedialog.askopenfilename(
            title="Select Baseline File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.baseline_file.set(filename)
            
    def browse_experiment(self):
        filename = filedialog.askopenfilename(
            title="Select Experiment File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.experiment_file.set(filename)
            
    def browse_output(self):
        dirname = filedialog.askdirectory(title="Select Output Directory")
        if dirname:
            self.output_dir.set(dirname)
            
    def log(self, message):
        """Add message to log with timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
        
    def open_output_folder(self):
        output_dir = self.output_dir.get()
        if os.path.exists(output_dir):
            os.startfile(output_dir)
        else:
            messagebox.showerror("Error", "Output directory does not exist")
            
    def run_analysis(self):
        """Run the analysis in a separate thread"""
        if self.mode.get() == "single":
            # Validate single mode inputs
            if not self.baseline_file.get():
                messagebox.showerror("Error", "Please select a baseline file")
                return
            if not self.experiment_file.get():
                messagebox.showerror("Error", "Please select an experiment file")
                return
            if not os.path.exists(self.baseline_file.get()):
                messagebox.showerror("Error", "Baseline file does not exist")
                return
            if not os.path.exists(self.experiment_file.get()):
                messagebox.showerror("Error", "Experiment file does not exist")
                return
        else:
            # Validate multi mode inputs
            if len(self.experiments_list) < 1:
                messagebox.showerror("Error", "Please add at least one experiment")
                return
            for exp in self.experiments_list:
                if not os.path.exists(exp['baseline']):
                    messagebox.showerror("Error", f"Baseline file not found: {exp['baseline']}")
                    return
                if not os.path.exists(exp['experiment']):
                    messagebox.showerror("Error", f"Experiment file not found: {exp['experiment']}")
                    return
            
        # Disable button and start progress
        self.run_button.config(state=tk.DISABLED)
        self.progress.start()
        self.status_var.set("⏳ Processing...")
        
        # Run in separate thread
        thread = threading.Thread(target=self._run_analysis_thread)
        thread.daemon = True
        thread.start()
        
    def _run_analysis_thread(self):
        """The actual analysis work"""
        try:
            self.log("=" * 60)
            self.log("Starting analysis...")
            self.log("=" * 60)
            
            # Get parameters
            pre_ignore = self.pre_treatment_ignore_minutes.get()
            post_ignore = self.post_treatment_ignore_minutes.get()
            pre_avg = self.pre_treatment_avg_seconds.get()
            post_avg = self.post_treatment_avg_seconds.get()
            
            self.log(f"Parameters:")
            self.log(f"  Pre-treatment: ignore first {pre_ignore} min, average every {pre_avg} sec")
            self.log(f"  Post-treatment: ignore first {post_ignore} min, average every {post_avg} sec")
            
            timestamp = datetime.now().strftime("%Y-%m-%d %H%M%S")
            output_dir = self.output_dir.get()
            
            if self.mode.get() == "single":
                self._run_single_analysis(timestamp, output_dir, pre_ignore, post_ignore, pre_avg, post_avg)
            else:
                self._run_multi_analysis(timestamp, output_dir, pre_ignore, post_ignore, pre_avg, post_avg)
            
            self.log("\n" + "=" * 60)
            self.log("Analysis complete!")
            self.log("=" * 60)
            
            self.root.after(0, lambda: messagebox.showinfo("Success", "Analysis completed successfully!"))
            
        except Exception as e:
            self.log(f"\nERROR: {str(e)}")
            self.log(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error", f"Analysis failed: {str(e)}"))
            
        finally:
            self.root.after(0, self._finish_analysis)
            
    def _run_single_analysis(self, timestamp, output_dir, pre_ignore, post_ignore, pre_avg, post_avg):
        """Run single experiment analysis"""
        # Process baseline file
        self.log("\nProcessing baseline file...")
        baseline_processed, baseline_grouped = self.process_excel_file(
            self.baseline_file.get(), pre_avg, post_avg, pre_ignore
        )
        
        # Save baseline files
        baseline_name = os.path.splitext(os.path.basename(self.baseline_file.get()))[0]
        baseline_processed_path = os.path.join(output_dir, f"{timestamp} {baseline_name}_processed.xlsx")
        baseline_grouped_path = os.path.join(output_dir, f"{timestamp} {baseline_name}_grouped.xlsx")
        
        self.save_excel(baseline_processed, baseline_processed_path)
        self.save_excel(baseline_grouped, baseline_grouped_path)
        self.log(f"  Saved: {os.path.basename(baseline_processed_path)}")
        self.log(f"  Saved: {os.path.basename(baseline_grouped_path)}")
        
        # Process experiment file
        self.log("\nProcessing experiment file...")
        exp_processed, exp_grouped = self.process_excel_file(
            self.experiment_file.get(), pre_avg, post_avg, pre_ignore
        )
        
        # Save experiment files
        exp_name = os.path.splitext(os.path.basename(self.experiment_file.get()))[0]
        exp_processed_path = os.path.join(output_dir, f"{timestamp} {exp_name}_processed.xlsx")
        exp_grouped_path = os.path.join(output_dir, f"{timestamp} {exp_name}_grouped.xlsx")
        
        self.save_excel(exp_processed, exp_processed_path)
        self.save_excel(exp_grouped, exp_grouped_path)
        self.log(f"  Saved: {os.path.basename(exp_processed_path)}")
        self.log(f"  Saved: {os.path.basename(exp_grouped_path)}")
        
        # Run analysis
        self.log("\nRunning analysis...")
        self.run_analysis_on_data(
            baseline_grouped, exp_grouped, baseline_processed, exp_processed,
            output_dir, timestamp, pre_ignore, post_ignore
        )
        
    def _run_multi_analysis(self, timestamp, output_dir, pre_ignore, post_ignore, pre_avg, post_avg):
        """Run multi-experiment comparison analysis"""
        all_experiment_data = {}
        
        for exp in self.experiments_list:
            exp_name = exp['name']
            self.log(f"\nProcessing {exp_name}...")
            
            # Process baseline
            self.log(f"  Processing baseline: {os.path.basename(exp['baseline'])}")
            baseline_processed, baseline_grouped = self.process_excel_file(
                exp['baseline'], pre_avg, post_avg, pre_ignore
            )
            
            # Process experiment
            self.log(f"  Processing experiment: {os.path.basename(exp['experiment'])}")
            exp_processed, exp_grouped = self.process_excel_file(
                exp['experiment'], pre_avg, post_avg, pre_ignore
            )
            
            # Save files
            bl_name = os.path.splitext(os.path.basename(exp['baseline']))[0]
            ex_name = os.path.splitext(os.path.basename(exp['experiment']))[0]
            
            self.save_excel(baseline_grouped, os.path.join(output_dir, f"{timestamp} {bl_name}_grouped.xlsx"))
            self.save_excel(exp_grouped, os.path.join(output_dir, f"{timestamp} {ex_name}_grouped.xlsx"))
            
            # Get common groups
            groups = sorted(set(baseline_grouped.keys()) & set(exp_grouped.keys()))
            self.log(f"  Groups found: {groups}")
            
            # Filter out empty groups
            valid_groups = []
            for g in groups:
                if len(baseline_grouped[g]) > 0 and len(exp_grouped[g]) > 0:
                    valid_groups.append(g)
                else:
                    self.log(f"  WARNING: Skipping {g} - no data")
            
            all_experiment_data[exp_name] = {
                'baseline': baseline_grouped,
                'experiment': exp_grouped,
                'groups': valid_groups,
            }
        
        # Create comparison plots
        self.log("\nCreating comparison plots...")
        self.create_multi_experiment_plots(all_experiment_data, output_dir, timestamp, post_ignore)
        
    def _finish_analysis(self):
        """Clean up after analysis"""
        self.run_button.config(state=tk.NORMAL)
        self.progress.stop()
        self.status_var.set("✓ Ready")
        
    def save_excel(self, sheets_dict, filepath):
        """Save dictionary of dataframes to Excel file"""
        if sheets_dict:
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                for sheet_name, df in sheets_dict.items():
                    safe_name = sheet_name[:31]
                    df.to_excel(writer, sheet_name=safe_name, index=False)
                    
    # ============ Processing Functions ============
    
    def is_sheet_empty(self, df):
        """Check if sheet only has header row"""
        if len(df) <= 1:
            return True
        if len(df) > 1:
            data_rows = df.iloc[1:]
            if data_rows.dropna(how='all').empty:
                return True
        return False
        
    def extract_group_letter(self, sheet_name):
        """Extract the group letter from sheet name"""
        match = re.match(r'\d+([a-zA-Z])\.', sheet_name)
        if match:
            return match.group(1).lower()
        return None
        
    def add_minutes_from_time_zero(self, df, col_b):
        """Add a column showing minutes before/after time 0"""
        if col_b is None or col_b not in df.columns:
            return df
            
        deliver_mask = df[col_b].astype(str).str.contains('Deliver Compound', case=False, na=False)
        antidote_mask = df[col_b].astype(str).str.contains('Measuring Antidote', case=False, na=False)
        
        deliver_indices = df.index[deliver_mask].tolist()
        antidote_indices = df.index[antidote_mask].tolist()
        
        if not deliver_indices:
            df['Minutes_from_Time0'] = None
            return df
            
        deliver_idx = deliver_indices[0]
        antidote_idx = antidote_indices[0] if antidote_indices else deliver_idx
        
        deliver_time = df.loc[deliver_idx, 'Time'] if 'Time' in df.columns else None
        antidote_time = df.loc[antidote_idx, 'Time'] if 'Time' in df.columns else None
        
        minutes_list = []
        
        for idx in df.index:
            row_time = df.loc[idx, 'Time'] if 'Time' in df.columns else None
            
            is_deliver = deliver_mask.loc[idx] if idx in deliver_mask.index else False
            is_antidote = antidote_mask.loc[idx] if idx in antidote_mask.index else False
            
            if is_deliver or is_antidote:
                minutes_list.append(0)
            elif row_time is not None and deliver_time is not None:
                try:
                    row_time_str = str(row_time)
                    deliver_time_str = str(deliver_time)
                    antidote_time_str = str(antidote_time) if antidote_time else deliver_time_str
                    
                    def time_to_minutes(t):
                        parts = t.split(':')
                        return int(parts[0]) * 60 + int(parts[1])
                    
                    row_minutes = time_to_minutes(row_time_str)
                    deliver_minutes = time_to_minutes(deliver_time_str)
                    antidote_minutes = time_to_minutes(antidote_time_str)
                    
                    if idx < deliver_idx:
                        diff = row_minutes - deliver_minutes
                        if diff == 0:
                            diff = idx - deliver_idx
                        minutes_list.append(diff)
                    elif idx > antidote_idx:
                        diff = row_minutes - antidote_minutes
                        minutes_list.append(diff)
                    else:
                        minutes_list.append(0)
                except:
                    minutes_list.append(None)
            else:
                minutes_list.append(None)
                
        df['Minutes_from_Time0'] = minutes_list
        return df
        
    def process_excel_file(self, input_path, pre_avg_seconds, post_avg_seconds, pre_ignore_minutes):
        """Process a single Excel file"""
        xl = pd.ExcelFile(input_path)
        sheet_names = xl.sheet_names
        self.log(f"  Found {len(sheet_names)} sheets")
        
        processed_sheets = {}
        
        for sheet_name in sheet_names:
            df = pd.read_excel(input_path, sheet_name=sheet_name)
            
            if self.is_sheet_empty(df):
                continue
                
            # Split date and time
            first_col = df.columns[0]
            try:
                if pd.api.types.is_datetime64_any_dtype(df[first_col]):
                    datetime_col = df[first_col]
                else:
                    datetime_col = pd.to_datetime(df[first_col], errors='coerce')
                    
                valid_count = datetime_col.notna().sum()
                if valid_count > 0:
                    df.insert(0, 'Date_New', datetime_col.dt.date)
                    df.insert(1, 'Time_New', datetime_col.dt.strftime('%H:%M:%S'))
                    original_col_name = df.columns[2]
                    df = df.drop(columns=[original_col_name])
                    df = df.rename(columns={'Date_New': 'Date', 'Time_New': 'Time'})
            except:
                pass
                
            # Delete rows between markers
            col_b = None
            if len(df.columns) > 2:
                col_b = df.columns[2]
                deliver_mask = df[col_b].astype(str).str.contains('Deliver Compound', case=False, na=False)
                antidote_mask = df[col_b].astype(str).str.contains('Measuring Antidote', case=False, na=False)
                
                deliver_indices = df.index[deliver_mask].tolist()
                antidote_indices = df.index[antidote_mask].tolist()
                marker_indices = deliver_indices + antidote_indices
                marker_set = set(marker_indices)
                
                rows_to_delete = []
                for deliver_idx in deliver_indices:
                    next_antidote = [idx for idx in antidote_indices if idx > deliver_idx]
                    if next_antidote:
                        antidote_idx = next_antidote[0]
                        for idx in df.index:
                            if idx > deliver_idx and idx < antidote_idx and idx not in marker_set:
                                rows_to_delete.append(idx)
                                
                rows_to_delete = list(set(rows_to_delete))
                if rows_to_delete:
                    df = df.drop(index=rows_to_delete)
                    
            df = df.reset_index(drop=True)
            
            # Add Minutes_from_Time0 first
            df = self.add_minutes_from_time_zero(df, col_b)
            
            # Average data based on pre/post settings
            if 'Time' in df.columns and 'Minutes_from_Time0' in df.columns:
                marker_mask = pd.Series([False] * len(df), index=df.index)
                if col_b is not None and col_b in df.columns:
                    marker_mask = (
                        df[col_b].astype(str).str.contains('Deliver Compound', case=False, na=False) |
                        df[col_b].astype(str).str.contains('Measuring Antidote', case=False, na=False)
                    )
                    
                df_markers = df[marker_mask].copy()
                df_data = df[~marker_mask].copy()
                
                df_pre = df_data[df_data['Minutes_from_Time0'] < 0].copy()
                df_post = df_data[df_data['Minutes_from_Time0'] >= 0].copy()
                
                numeric_cols = df_data.select_dtypes(include=['number']).columns.tolist()
                if 'Minutes_from_Time0' in numeric_cols:
                    numeric_cols.remove('Minutes_from_Time0')
                    
                averaged_parts = []
                
                # Average pre-treatment data
                if len(df_pre) > 0 and numeric_cols:
                    if pre_avg_seconds >= 60:
                        df_pre['TimeGroup'] = df_pre['Time'].astype(str).str[:5]
                    else:
                        df_pre['TimeGroup'] = self.round_time_to_seconds(df_pre['Time'], pre_avg_seconds)
                        
                    group_cols = []
                    if 'Date' in df_pre.columns:
                        group_cols.append('Date')
                    group_cols.append('TimeGroup')
                    
                    agg_dict = {col: 'mean' for col in numeric_cols}
                    non_numeric_cols = [col for col in df_pre.columns 
                                       if col not in numeric_cols 
                                       and col not in group_cols 
                                       and col != 'Time']
                    for col in non_numeric_cols:
                        agg_dict[col] = 'first'
                        
                    df_pre_avg = df_pre.groupby(group_cols, as_index=False).agg(agg_dict)
                    df_pre_avg = df_pre_avg.rename(columns={'TimeGroup': 'Time'})
                    averaged_parts.append(df_pre_avg)
                    
                # Average post-treatment data
                if len(df_post) > 0 and numeric_cols:
                    df_post['TimeGroup'] = self.round_time_to_seconds(df_post['Time'], post_avg_seconds)
                    
                    group_cols = []
                    if 'Date' in df_post.columns:
                        group_cols.append('Date')
                    group_cols.append('TimeGroup')
                    
                    agg_dict = {col: 'mean' for col in numeric_cols}
                    non_numeric_cols = [col for col in df_post.columns 
                                       if col not in numeric_cols 
                                       and col not in group_cols 
                                       and col != 'Time']
                    for col in non_numeric_cols:
                        agg_dict[col] = 'first'
                        
                    df_post_avg = df_post.groupby(group_cols, as_index=False).agg(agg_dict)
                    df_post_avg = df_post_avg.rename(columns={'TimeGroup': 'Time'})
                    averaged_parts.append(df_post_avg)
                    
                if averaged_parts:
                    df_averaged = pd.concat(averaged_parts, ignore_index=True)
                    
                    if len(df_markers) > 0:
                        df_markers = df_markers.copy()
                        df_markers['Time'] = df_markers['Time'].astype(str).str[:8]
                        
                    df = pd.concat([df_averaged, df_markers], ignore_index=True)
                    
                    sort_cols = []
                    if 'Date' in df.columns:
                        sort_cols.append('Date')
                    if 'Time' in df.columns:
                        sort_cols.append('Time')
                    if sort_cols:
                        df = df.sort_values(by=sort_cols).reset_index(drop=True)
                        
            # Recalculate Minutes_from_Time0
            df = self.add_minutes_from_time_zero(df, col_b)
            
            processed_sheets[sheet_name] = df
            
        # Group sheets by letter
        grouped_sheets = self.group_sheets_by_letter(processed_sheets)
        
        return processed_sheets, grouped_sheets
        
    def round_time_to_seconds(self, time_series, seconds):
        """Round time to nearest X seconds"""
        def round_single(time_str):
            try:
                parts = str(time_str).split(':')
                if len(parts) >= 2:
                    h = int(parts[0])
                    m = int(parts[1])
                    s = int(parts[2]) if len(parts) >= 3 else 0
                    total_sec = h * 3600 + m * 60 + s
                    rounded_sec = (total_sec // seconds) * seconds
                    h_new = rounded_sec // 3600
                    m_new = (rounded_sec % 3600) // 60
                    s_new = rounded_sec % 60
                    return f"{h_new:02d}:{m_new:02d}:{s_new:02d}"
            except:
                pass
            return str(time_str)[:8]
        return time_series.apply(round_single)
        
    def group_sheets_by_letter(self, processed_sheets):
        """Group sheets by their letter and average"""
        groups = {}
        for sheet_name, df in processed_sheets.items():
            letter = self.extract_group_letter(sheet_name)
            if letter:
                if letter not in groups:
                    groups[letter] = []
                groups[letter].append((sheet_name, df))
                
        grouped_sheets = {}
        
        for letter, sheet_list in sorted(groups.items()):
            dfs = [s[1].copy() for s in sheet_list]
            
            if not all('Minutes_from_Time0' in df.columns for df in dfs):
                continue
                
            all_minutes = set()
            for df in dfs:
                all_minutes.update(df['Minutes_from_Time0'].dropna().unique())
            all_minutes = sorted(all_minutes)
            
            if len(all_minutes) == 0:
                grouped_sheets[f"group {letter}"] = pd.DataFrame()
                continue
                
            numeric_cols = []
            for df in dfs:
                for col in df.select_dtypes(include=['number']).columns:
                    if col not in numeric_cols and col != 'Minutes_from_Time0':
                        numeric_cols.append(col)
                        
            non_numeric_cols = [col for col in dfs[0].columns 
                              if col not in numeric_cols 
                              and col != 'Minutes_from_Time0']
            
            result_data = []
            
            for minute in all_minutes:
                row_data = {'Minutes_from_Time0': minute}
                
                minute_rows = []
                for df in dfs:
                    matching_rows = df[df['Minutes_from_Time0'] == minute]
                    if len(matching_rows) > 0:
                        minute_rows.append(matching_rows.iloc[0])
                        
                if minute_rows:
                    for col in numeric_cols:
                        values = []
                        for row in minute_rows:
                            if col in row.index and pd.notna(row[col]):
                                values.append(row[col])
                        if values:
                            row_data[col] = sum(values) / len(values)
                        else:
                            row_data[col] = None
                            
                    for col in non_numeric_cols:
                        for row in minute_rows:
                            if col in row.index and pd.notna(row[col]):
                                row_data[col] = row[col]
                                break
                        else:
                            row_data[col] = None
                            
                result_data.append(row_data)
                
            grouped_df = pd.DataFrame(result_data)
            
            cols_order = ['Minutes_from_Time0']
            for col in non_numeric_cols:
                if col in grouped_df.columns and col not in cols_order:
                    cols_order.append(col)
            for col in numeric_cols:
                if col in grouped_df.columns and col not in cols_order:
                    cols_order.append(col)
            for col in grouped_df.columns:
                if col not in cols_order:
                    cols_order.append(col)
                    
            grouped_df = grouped_df[[c for c in cols_order if c in grouped_df.columns]]
            
            group_name = f"group {letter}"
            grouped_sheets[group_name] = grouped_df
            
        return grouped_sheets
        
    # ============ Analysis Functions ============
    
    def run_analysis_on_data(self, baseline_grouped, exp_grouped, baseline_processed, exp_processed,
                             output_dir, timestamp, pre_ignore, post_ignore):
        """Run the analysis on processed data"""
        
        # Get common groups
        groups = sorted(set(baseline_grouped.keys()) & set(exp_grouped.keys()))
        
        # Filter out empty groups
        valid_groups = []
        for g in groups:
            if len(baseline_grouped[g]) > 0 and len(exp_grouped[g]) > 0:
                valid_groups.append(g)
            else:
                self.log(f"  WARNING: Skipping {g} - baseline or experiment data is empty")
        groups = valid_groups
        
        self.log(f"  Valid groups: {groups}")
        
        if not groups:
            self.log("  WARNING: No valid groups found!")
            return
            
        key_columns = ['f', 'TVb', 'MVb', 'Penh', 'Ti', 'Te', 'PIFb', 'PEFb']
        
        # Create plots directory
        plots_dir = os.path.join(output_dir, "analysis_plots")
        os.makedirs(plots_dir, exist_ok=True)
        
        # Calculate statistics
        all_stats = {}
        comparison_results = {}
        
        for group in groups:
            all_stats[group] = {}
            comparison_results[group] = {}
            
            baseline = baseline_grouped[group]
            experiment = exp_grouped[group]
            
            # Get baseline data (post-treatment, ignoring first X minutes)
            if 'Minutes_from_Time0' in baseline.columns:
                baseline_filtered = baseline[baseline['Minutes_from_Time0'] >= post_ignore].copy()
            else:
                baseline_filtered = baseline.copy()
                
            # Get experiment pre and post
            if 'Minutes_from_Time0' in experiment.columns:
                exp_pre = experiment[experiment['Minutes_from_Time0'] < -pre_ignore].copy()
                exp_post = experiment[experiment['Minutes_from_Time0'] >= post_ignore].copy()
            else:
                exp_pre = pd.DataFrame()
                exp_post = experiment.copy()
                
            for col in key_columns:
                bl_vals = baseline_filtered[col].dropna() if col in baseline_filtered.columns else pd.Series([])
                pre_vals = exp_pre[col].dropna() if col in exp_pre.columns else pd.Series([])
                post_vals = exp_post[col].dropna() if col in exp_post.columns else pd.Series([])
                
                all_stats[group][col] = {
                    'baseline': {'mean': bl_vals.mean(), 'std': bl_vals.std(), 
                                'n': len(bl_vals), 'sem': bl_vals.std() / np.sqrt(len(bl_vals)) if len(bl_vals) > 0 else 0,
                                'values': bl_vals} if len(bl_vals) > 0 else None,
                    'exp_pre': {'mean': pre_vals.mean(), 'std': pre_vals.std(),
                               'n': len(pre_vals), 'sem': pre_vals.std() / np.sqrt(len(pre_vals)) if len(pre_vals) > 0 else 0,
                               'values': pre_vals} if len(pre_vals) > 0 else None,
                    'exp_post': {'mean': post_vals.mean(), 'std': post_vals.std(),
                                'n': len(post_vals), 'sem': post_vals.std() / np.sqrt(len(post_vals)) if len(post_vals) > 0 else 0,
                                'values': post_vals} if len(post_vals) > 0 else None,
                }
                
                # Statistical tests
                if len(bl_vals) > 0 and len(pre_vals) > 0:
                    t, p = stats.ttest_ind(bl_vals, pre_vals)
                    diff = ((pre_vals.mean() - bl_vals.mean()) / bl_vals.mean() * 100) if bl_vals.mean() != 0 else 0
                    sig = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else ""
                    comparison_results[group][f'{col}_pre_vs_bl'] = {'p': p, 'diff': diff, 'sig': sig}
                    
                if len(bl_vals) > 0 and len(post_vals) > 0:
                    t, p = stats.ttest_ind(bl_vals, post_vals)
                    diff = ((post_vals.mean() - bl_vals.mean()) / bl_vals.mean() * 100) if bl_vals.mean() != 0 else 0
                    sig = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else ""
                    comparison_results[group][f'{col}_post_vs_bl'] = {'p': p, 'diff': diff, 'sig': sig}
                    
                if len(pre_vals) > 0 and len(post_vals) > 0:
                    t, p = stats.ttest_ind(pre_vals, post_vals)
                    diff = ((post_vals.mean() - pre_vals.mean()) / pre_vals.mean() * 100) if pre_vals.mean() != 0 else 0
                    sig = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else ""
                    comparison_results[group][f'{col}_treatment'] = {'p': p, 'diff': diff, 'sig': sig}
                    
        # Save statistical summary
        stats_path = os.path.join(output_dir, f"{timestamp} statistical_analysis_summary.txt")
        with open(stats_path, 'w', encoding='utf-8') as f:
            f.write("STATISTICAL ANALYSIS SUMMARY\n")
            f.write("=" * 70 + "\n")
            f.write(f"Analysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Pre-treatment: ignore first {pre_ignore} min\n")
            f.write(f"Post-treatment: ignore first {post_ignore} min\n\n")
            
            for group in groups:
                f.write(f"\n{'='*70}\n")
                f.write(f"GROUP: {group.upper()}\n")
                f.write(f"{'='*70}\n\n")
                
                for comparison_type, title in [('_pre_vs_bl', 'PRE vs BASELINE'), 
                                                ('_post_vs_bl', 'POST vs BASELINE'),
                                                ('_treatment', 'TREATMENT EFFECT')]:
                    f.write(f"{title}\n")
                    f.write("-" * 50 + "\n")
                    for col in key_columns:
                        key = f'{col}{comparison_type}'
                        if key in comparison_results[group]:
                            r = comparison_results[group][key]
                            f.write(f"  {col}: Diff={r['diff']:+.1f}%, p={r['p']:.4f} {r['sig']}\n")
                    f.write("\n")
                    
        self.log(f"  Saved: {os.path.basename(stats_path)}")
        
        # Create visualizations
        self.create_timecourse_plots(baseline_grouped, exp_grouped, groups, key_columns,
                                     plots_dir, timestamp, post_ignore)
        self.create_comparison_plots(all_stats, groups, key_columns, plots_dir, timestamp)
        
        self.log(f"  All plots saved to: {plots_dir}")
        
    def create_timecourse_plots(self, baseline_grouped, exp_grouped, groups, key_columns,
                                plots_dir, timestamp, post_ignore):
        """Create time-course plots"""
        
        group_colors = {
            'group a': '#E74C3C', 'group b': '#3498DB', 'group c': '#27AE60',
            'group d': '#9B59B6', 'group e': '#F39C12', 'group f': '#1ABC9C',
        }
        group_markers = {
            'group a': 'o', 'group b': 's', 'group c': '^',
            'group d': 'D', 'group e': 'v', 'group f': 'p',
        }
        
        param_sets = [
            (key_columns[:4], 'timecourse_1.png', 'Time Course (Part 1)'),
            (key_columns[4:], 'timecourse_2.png', 'Time Course (Part 2)'),
        ]
        
        for param_set, filename_suffix, title in param_sets:
            fig, axes = plt.subplots(2, 2, figsize=(16, 12))
            axes = axes.flatten()
            
            for j, col in enumerate(param_set):
                ax = axes[j]
                
                for group in groups:
                    color = group_colors.get(group, 'gray')
                    marker = group_markers.get(group, 'o')
                    
                    baseline = baseline_grouped[group]
                    experiment = exp_grouped[group]
                    
                    all_times = []
                    all_means = []
                    
                    # Baseline data
                    if 'Minutes_from_Time0' in baseline.columns and col in baseline.columns:
                        bl_data = baseline[baseline['Minutes_from_Time0'] >= post_ignore].copy()
                        if len(bl_data) > 0:
                            max_time = bl_data['Minutes_from_Time0'].max()
                            last_10 = bl_data[bl_data['Minutes_from_Time0'] >= max_time - 10].copy()
                            if len(last_10) > 0:
                                orig_min = last_10['Minutes_from_Time0'].min()
                                orig_max = last_10['Minutes_from_Time0'].max()
                                if orig_max > orig_min:
                                    last_10['mapped'] = -20 + (last_10['Minutes_from_Time0'] - orig_min) / (orig_max - orig_min) * 10
                                else:
                                    last_10['mapped'] = -15
                                last_10['mapped_int'] = last_10['mapped'].round().astype(int)
                                for minute, grp in last_10.groupby('mapped_int'):
                                    vals = grp[col].dropna()
                                    if len(vals) > 0:
                                        all_times.append(minute)
                                        all_means.append(vals.mean())
                                        
                    # Experiment data
                    if 'Minutes_from_Time0' in experiment.columns and col in experiment.columns:
                        exp_data = experiment[experiment['Minutes_from_Time0'] <= 30].copy()
                        exp_data['minute'] = exp_data['Minutes_from_Time0'].round().astype(int)
                        for minute, grp in exp_data.groupby('minute'):
                            vals = grp[col].dropna()
                            if len(vals) > 0:
                                all_times.append(minute)
                                all_means.append(vals.mean())
                                
                    if all_times:
                        sorted_idx = np.argsort(all_times)
                        all_times = np.array(all_times)[sorted_idx]
                        all_means = np.array(all_means)[sorted_idx]
                        ax.plot(all_times, all_means, f'-{marker}', color=color,
                               label=group.upper(), markersize=5, linewidth=1.5)
                               
                ax.axvline(x=0, color='black', linestyle='--', linewidth=2)
                ax.set_xlim(-25, 32)
                # Add shading: blue (-20 to -10), green (-10 to 0), red (0 to 30)
                ax.axvspan(-20, -10, alpha=0.2, color='blue')
                ax.axvspan(-10, 0, alpha=0.2, color='green')
                ax.axvspan(0, 30, alpha=0.2, color='red')
                ax.set_xlabel('Minutes from Treatment')
                ax.set_ylabel(col)
                ax.set_title(f'{col} - Time Course')
                if j == 0:
                    ax.legend(loc='upper right', fontsize=8)
                    
            plt.suptitle(title, fontsize=14, fontweight='bold')
            plt.tight_layout()
            plt.savefig(os.path.join(plots_dir, f'{timestamp} {filename_suffix}'), dpi=150)
            plt.close()
            gc.collect()
            self.log(f"  Saved: {timestamp} {filename_suffix}")
            
    def create_comparison_plots(self, all_stats, groups, key_columns, plots_dir, timestamp):
        """Create comparison bar plots"""
        
        for group in groups:
            fig, axes = plt.subplots(2, 4, figsize=(16, 8))
            axes = axes.flatten()
            
            for idx, col in enumerate(key_columns):
                ax = axes[idx]
                col_stats = all_stats[group][col]
                
                means = [
                    col_stats['baseline']['mean'] if col_stats['baseline'] else 0,
                    col_stats['exp_pre']['mean'] if col_stats['exp_pre'] else 0,
                    col_stats['exp_post']['mean'] if col_stats['exp_post'] else 0
                ]
                sems = [
                    col_stats['baseline']['sem'] if col_stats['baseline'] else 0,
                    col_stats['exp_pre']['sem'] if col_stats['exp_pre'] else 0,
                    col_stats['exp_post']['sem'] if col_stats['exp_post'] else 0
                ]
                
                colors = ['steelblue', 'lightsalmon', 'coral']
                ax.bar([0, 1, 2], means, yerr=sems, capsize=4, color=colors, edgecolor='black')
                ax.set_ylabel(col)
                ax.set_title(col)
                ax.set_xticks([0, 1, 2])
                ax.set_xticklabels(['Baseline', 'Exp Pre', 'Exp Post'], rotation=15)
                
            plt.suptitle(f'{group.upper()}: Comparison', fontsize=14, fontweight='bold')
            plt.tight_layout()
            filename = f'{timestamp} comparison_{group.replace(" ", "_")}.png'
            plt.savefig(os.path.join(plots_dir, filename), dpi=150)
            plt.close()
            gc.collect()
            self.log(f"  Saved: {filename}")
            
    def create_multi_experiment_plots(self, all_experiment_data, output_dir, timestamp, post_ignore):
        """Create plots comparing multiple experiments"""
        
        plots_dir = os.path.join(output_dir, "comparison_plots")
        os.makedirs(plots_dir, exist_ok=True)
        
        key_columns = ['f', 'TVb', 'MVb', 'Penh', 'Ti', 'Te', 'PIFb', 'PEFb']
        
        # Color palettes for experiments
        exp_colors = [
            {'group a': '#E74C3C', 'group b': '#3498DB', 'group c': '#27AE60', 
             'group d': '#9B59B6', 'group e': '#F39C12', 'group f': '#1ABC9C'},
            {'group a': '#C0392B', 'group b': '#2980B9', 'group c': '#1E8449',
             'group d': '#7D3C98', 'group e': '#D68910', 'group f': '#148F77'},
            {'group a': '#A93226', 'group b': '#1F618D', 'group c': '#145A32',
             'group d': '#6C3483', 'group e': '#B9770E', 'group f': '#117864'},
        ]
        exp_markers = ['o', 's', '^', 'D', 'v', 'p']
        exp_linestyles = ['-', '--', '-.', ':']
        
        param_sets = [
            (key_columns[:4], 'combined_timecourse_1.png', 'Multi-Experiment Comparison (Part 1)'),
            (key_columns[4:], 'combined_timecourse_2.png', 'Multi-Experiment Comparison (Part 2)'),
        ]
        
        for param_set, filename_suffix, title in param_sets:
            fig, axes = plt.subplots(2, 2, figsize=(16, 12))
            axes = axes.flatten()
            
            legend_handles = []
            legend_labels = []
            
            for j, col in enumerate(param_set):
                ax = axes[j]
                
                for exp_idx, (exp_name, exp_data) in enumerate(all_experiment_data.items()):
                    baseline_grouped = exp_data['baseline']
                    exp_grouped = exp_data['experiment']
                    groups = exp_data['groups']
                    
                    colors = exp_colors[exp_idx % len(exp_colors)]
                    marker = exp_markers[exp_idx % len(exp_markers)]
                    linestyle = exp_linestyles[exp_idx % len(exp_linestyles)]
                    
                    for group in groups:
                        color = colors.get(group, 'gray')
                        
                        baseline = baseline_grouped.get(group, pd.DataFrame())
                        experiment = exp_grouped.get(group, pd.DataFrame())
                        
                        if len(baseline) == 0 or len(experiment) == 0:
                            continue
                        
                        all_times = []
                        all_means = []
                        
                        # Baseline data
                        if 'Minutes_from_Time0' in baseline.columns and col in baseline.columns:
                            bl_data = baseline[baseline['Minutes_from_Time0'] >= post_ignore].copy()
                            if len(bl_data) > 0:
                                max_time = bl_data['Minutes_from_Time0'].max()
                                last_10 = bl_data[bl_data['Minutes_from_Time0'] >= max_time - 10].copy()
                                if len(last_10) > 0:
                                    orig_min = last_10['Minutes_from_Time0'].min()
                                    orig_max = last_10['Minutes_from_Time0'].max()
                                    if orig_max > orig_min:
                                        last_10['mapped'] = -20 + (last_10['Minutes_from_Time0'] - orig_min) / (orig_max - orig_min) * 10
                                    else:
                                        last_10['mapped'] = -15
                                    last_10['mapped_int'] = last_10['mapped'].round().astype(int)
                                    for minute, grp in last_10.groupby('mapped_int'):
                                        vals = grp[col].dropna()
                                        if len(vals) > 0:
                                            all_times.append(minute)
                                            all_means.append(vals.mean())
                                            
                        # Experiment data
                        if 'Minutes_from_Time0' in experiment.columns and col in experiment.columns:
                            exp_df = experiment[experiment['Minutes_from_Time0'] <= 30].copy()
                            exp_df['minute'] = exp_df['Minutes_from_Time0'].round().astype(int)
                            for minute, grp in exp_df.groupby('minute'):
                                vals = grp[col].dropna()
                                if len(vals) > 0:
                                    all_times.append(minute)
                                    all_means.append(vals.mean())
                                    
                        if all_times:
                            sorted_idx = np.argsort(all_times)
                            all_times = np.array(all_times)[sorted_idx]
                            all_means = np.array(all_means)[sorted_idx]
                            
                            label = f"{exp_name} - {group.upper()}"
                            line, = ax.plot(all_times, all_means, f'{linestyle}{marker}', color=color,
                                          label=label, markersize=4, linewidth=1.5)
                            
                            if j == 0:
                                legend_handles.append(line)
                                legend_labels.append(label)
                                
                ax.axvline(x=0, color='black', linestyle='--', linewidth=2)
                ax.set_xlim(-25, 32)
                # Add shading: blue (-20 to -10), green (-10 to 0), red (0 to 30)
                ax.axvspan(-20, -10, alpha=0.2, color='blue')
                ax.axvspan(-10, 0, alpha=0.2, color='green')
                ax.axvspan(0, 30, alpha=0.2, color='red')
                ax.set_xlabel('Minutes from Treatment')
                ax.set_ylabel(col)
                ax.set_title(f'{col} - All Experiments')
                
            fig.legend(legend_handles, legend_labels, loc='center right', 
                      bbox_to_anchor=(1.15, 0.5), fontsize=8)
            plt.suptitle(title, fontsize=14, fontweight='bold')
            plt.tight_layout()
            plt.subplots_adjust(right=0.85)
            
            plt.savefig(os.path.join(plots_dir, f'{timestamp} {filename_suffix}'), dpi=150, bbox_inches='tight')
            plt.close()
            gc.collect()
            self.log(f"  Saved: {timestamp} {filename_suffix}")
            
        self.log(f"  Comparison plots saved to: {plots_dir}")


class AddExperimentDialog:
    """Dialog for adding a new experiment"""
    def __init__(self, parent):
        self.result = None
        
        self.top = tk.Toplevel(parent)
        self.top.title("Add New Experiment")
        self.top.geometry("550x220")
        self.top.transient(parent)
        self.top.grab_set()
        self.top.configure(bg='#F5F5F5')
        
        # Center on parent
        self.top.geometry(f"+{parent.winfo_x() + 100}+{parent.winfo_y() + 100}")
        
        frame = tk.Frame(self.top, bg='#F5F5F5', padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Experiment name
        ttk.Label(frame, text="Experiment Name:", 
                 font=('Segoe UI', 9, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=8)
        self.name_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.name_var, width=42).grid(row=0, column=1, padx=8, pady=8, columnspan=2, sticky=tk.EW)
        
        # Baseline file
        ttk.Label(frame, text="Baseline File:", 
                 font=('Segoe UI', 9, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=8)
        self.baseline_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.baseline_var, width=35).grid(row=1, column=1, padx=8, pady=8, sticky=tk.EW)
        ttk.Button(frame, text="📂 Browse", command=self.browse_baseline).grid(row=1, column=2, padx=5)
        
        # Experiment file
        ttk.Label(frame, text="Experiment File:", 
                 font=('Segoe UI', 9, 'bold')).grid(row=2, column=0, sticky=tk.W, pady=8)
        self.experiment_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.experiment_var, width=35).grid(row=2, column=1, padx=8, pady=8, sticky=tk.EW)
        ttk.Button(frame, text="📂 Browse", command=self.browse_experiment).grid(row=2, column=2, padx=5)
        
        frame.columnconfigure(1, weight=1)
        
        # Buttons
        btn_frame = tk.Frame(frame, bg='#F5F5F5')
        btn_frame.grid(row=3, column=0, columnspan=3, pady=(15, 0))
        
        ttk.Button(btn_frame, text="✓ Add", command=self.add, 
                  style='Primary.TButton').pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frame, text="✗ Cancel", command=self.cancel,
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=8)
        
    def browse_baseline(self):
        filename = filedialog.askopenfilename(
            title="Select Baseline File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.baseline_var.set(filename)
            if not self.name_var.get():
                # Auto-fill name from filename
                name = os.path.splitext(os.path.basename(filename))[0]
                self.name_var.set(name.replace('_', ' '))
                
    def browse_experiment(self):
        filename = filedialog.askopenfilename(
            title="Select Experiment File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.experiment_var.set(filename)
            
    def add(self):
        if not self.name_var.get():
            messagebox.showerror("Error", "Please enter an experiment name")
            return
        if not self.baseline_var.get():
            messagebox.showerror("Error", "Please select a baseline file")
            return
        if not self.experiment_var.get():
            messagebox.showerror("Error", "Please select an experiment file")
            return
            
        self.result = {
            'name': self.name_var.get(),
            'baseline': self.baseline_var.get(),
            'experiment': self.experiment_var.get(),
        }
        self.top.destroy()
        
    def cancel(self):
        self.top.destroy()


def main():
    root = tk.Tk()
    
    # Set theme
    style = ttk.Style()
    try:
        style.theme_use('vista')
    except:
        try:
            style.theme_use('clam')
        except:
            pass
    
    # Set window icon (if available)
    try:
        root.iconbitmap('icon.ico')
    except: pass
            
    app = ExperimentAnalyzerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
