import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pathlib import Path
from datetime import datetime
import re

class ModernButton(tk.Canvas):
    """Custom gradient button with rounded corners"""
    def __init__(self, parent, text, command, width=200, height=45, 
                 gradient_colors=None, text_color="black", **kwargs):
        super().__init__(parent, width=width, height=height, 
                        highlightthickness=0, **kwargs)
        
        self.command = command
        self.text = text
        self.width = width
        self.height = height
        self.gradient_colors = gradient_colors or ["#60a5fa", "#3b82f6"]
        self.text_color = text_color
        
        self.draw_button()
        self.bind("<Button-1>", lambda e: self.on_click())
        self.bind("<Enter>", lambda e: self.on_hover())
        self.bind("<Leave>", lambda e: self.on_leave())
        
    def draw_button(self, hover=False):
        self.delete("all")
        
        # Draw gradient background
        colors = self.gradient_colors if not hover else [self.lighten_color(c) for c in self.gradient_colors]
        steps = 20
        for i in range(steps):
            y1 = i * (self.height / steps)
            y2 = (i + 1) * (self.height / steps)
            color = self.interpolate_color(colors[0], colors[1], i / steps)
            self.create_rectangle(0, y1, self.width, y2, fill=color, outline="")
        
        # Draw rounded rectangle overlay for rounded effect
        self.create_oval(0, 0, 20, 20, fill=colors[0], outline="")
        self.create_oval(self.width-20, 0, self.width, 20, fill=colors[0], outline="")
        self.create_oval(0, self.height-20, 20, self.height, fill=colors[1], outline="")
        self.create_oval(self.width-20, self.height-20, self.width, self.height, fill=colors[1], outline="")
        
        # Draw text
        self.create_text(self.width/2, self.height/2, text=self.text, 
                        font=("Segoe UI", 11, "bold"), fill=self.text_color)
    
    def lighten_color(self, color):
        """Lighten a hex color"""
        color = color.lstrip('#')
        r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
        r = min(255, r + 20)
        g = min(255, g + 20)
        b = min(255, b + 20)
        return f'#{r:02x}{g:02x}{b:02x}'
    
    def interpolate_color(self, color1, color2, ratio):
        """Interpolate between two colors"""
        c1 = color1.lstrip('#')
        c2 = color2.lstrip('#')
        r1, g1, b1 = int(c1[0:2], 16), int(c1[2:4], 16), int(c1[4:6], 16)
        r2, g2, b2 = int(c2[0:2], 16), int(c2[2:4], 16), int(c2[4:6], 16)
        
        r = int(r1 + (r2 - r1) * ratio)
        g = int(g1 + (g2 - g1) * ratio)
        b = int(b1 + (b2 - b1) * ratio)
        return f'#{r:02x}{g:02x}{b:02x}'
    
    def on_hover(self):
        self.draw_button(hover=True)
        self.config(cursor="hand2")
    
    def on_leave(self):
        self.draw_button(hover=False)
        self.config(cursor="")
    
    def on_click(self):
        if self.command:
            self.command()

class ExcelMatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Matcher & Merger Pro")
        self.root.geometry("1450x900")
        
        self.primary_file = None
        self.ref_file = None
        self.primary_df = None
        self.ref_df = None
        self.column_mappings = []
        
        # Modern Colors with section backgrounds
        self.colors = {
            'primary': '#3b82f6',
            'primary_hover': '#2563eb',
            'success': '#10b981',
            'success_hover': '#059669',
            'danger': '#ef4444',
            'danger_hover': '#dc2626',
            'warning': '#f59e0b',
            'warning_hover': '#d97706',
            'secondary': '#6b7280',
            'bg_light': '#f0f4f8',
            'bg_card': '#ffffff',
            'text_primary': '#111827',
            'text_secondary': '#6b7280',
            'border': '#e5e7eb',
            # Section backgrounds
            'section_file': '#e0f2fe',      # Light blue
            'section_mapping': '#fef3c7',   # Light amber
            'section_info': '#dbeafe',      # Light sky blue
            'section_progress': '#d1fae5',  # Light green
            'section_log': '#fce7f3'        # Light pink
        }
        
        self.create_widgets()
    
    def log_message(self, message, level="info"):
        """Add message to log with timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        color_map = {
            "info": self.colors['text_primary'],
            "success": self.colors['success'],
            "warning": self.colors['warning'],
            "error": self.colors['danger']
        }
        
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"[{timestamp}] ", "timestamp")
        self.log_text.insert("end", f"{message}\n", level)
        self.log_text.tag_config("timestamp", foreground=self.colors['secondary'], font=("Segoe UI", 10))
        self.log_text.tag_config(level, foreground=color_map.get(level, self.colors['text_primary']))
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self.root.update()
    
    def normalize_value(self, val):
        """Normalize values for comparison"""
        if pd.isna(val) or val is None:
            return ""
        
        if isinstance(val, (pd.Timestamp, datetime)):
            return val.strftime('%Y-%m-%d').lower().strip()
        
        val_str = str(val).strip().lower()
        val_str = re.sub(r'\s+', ' ', val_str)
        val_str = re.sub(r'[,.\-_]', '', val_str)
        
        return val_str
    
    def create_widgets(self):
        # Main container
        main_frame = tk.Frame(self.root, bg=self.colors['bg_light'])
        main_frame.pack(fill="both", expand=True)
        
        # Header with gradient effect
        header = tk.Frame(main_frame, bg=self.colors['primary'], height=80)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        title = tk.Label(header, text="📊 Excel Data Matcher & Merger Pro", 
                        font=("Segoe UI", 22, "bold"), fg="white", bg=self.colors['primary'])
        title.pack(pady=25)
        
        # Main content container
        content = tk.Frame(main_frame, bg=self.colors['bg_light'])
        content.pack(fill="both", expand=True, padx=25, pady=25)
        
        # Left side (Main operations)
        left_frame = tk.Frame(content, bg=self.colors['bg_light'])
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 12))
        
        # Right side (Info, Progress, Logs)
        right_frame = tk.Frame(content, bg=self.colors['bg_light'], width=470)
        right_frame.pack(side="right", fill="both", padx=(12, 0))
        right_frame.pack_propagate(False)
        
        # === LEFT SIDE CONTENT ===
        
        # File Selection with light blue background
        file_frame = tk.LabelFrame(left_frame, text="  📁 Step 1: Select Excel Files  ", 
                                   font=("Segoe UI", 13, "bold"), bg=self.colors['section_file'], 
                                   fg=self.colors['text_primary'], padx=25, pady=25, 
                                   relief="solid", bd=2, labelanchor="n")
        file_frame.pack(fill="x", pady=(0, 18))
        
        # Primary File
        primary_frame = tk.Frame(file_frame, bg=self.colors['section_file'])
        primary_frame.pack(fill="x", pady=12)
        
        tk.Label(primary_frame, text="Primary File:", font=("Segoe UI", 11, "bold"), 
                bg=self.colors['section_file'], fg=self.colors['text_primary'], 
                width=15, anchor="w").pack(side="left")
        
        self.primary_label = tk.Label(primary_frame, text="No file selected", 
                                 fg=self.colors['text_secondary'], 
                                 bg=self.colors['section_file'], anchor="w", 
                                 font=("Segoe UI", 11))
        self.primary_label.pack(side="left", fill="x", expand=True, padx=12)
        
        ModernButton(primary_frame, "Browse Files", self.select_primary_file,
                    width=150, height=40, 
                    gradient_colors=["#60a5fa", "#3b82f6"],
                    text_color="black", bg=self.colors['section_file']).pack(side="right")
        
        # Reference File
        ref_frame = tk.Frame(file_frame, bg=self.colors['section_file'])
        ref_frame.pack(fill="x", pady=12)
        
        tk.Label(ref_frame, text="Reference File:", font=("Segoe UI", 11, "bold"), 
                bg=self.colors['section_file'], fg=self.colors['text_primary'], 
                width=15, anchor="w").pack(side="left")
        
        self.ref_label = tk.Label(ref_frame, text="No file selected", 
                                 fg=self.colors['text_secondary'], 
                                 bg=self.colors['section_file'], anchor="w", 
                                 font=("Segoe UI", 11))
        self.ref_label.pack(side="left", fill="x", expand=True, padx=12)
        
        ModernButton(ref_frame, "Browse Files", self.select_ref_file,
                    width=150, height=40,
                    gradient_colors=["#60a5fa", "#3b82f6"],
                    text_color="black", bg=self.colors['section_file']).pack(side="right")
        
        # Column Mapping with light amber background
        self.mapping_frame = tk.LabelFrame(left_frame, 
                                          text="  🔗 Step 2: Configure Column Matching  ", 
                                          font=("Segoe UI", 13, "bold"), 
                                          bg=self.colors['section_mapping'], 
                                          fg=self.colors['text_primary'], 
                                          padx=25, pady=25, relief="solid", bd=2,
                                          labelanchor="n")
        self.mapping_frame.pack(fill="both", expand=True, pady=(0, 18))
        
        mapping_controls = tk.Frame(self.mapping_frame, bg=self.colors['section_mapping'])
        mapping_controls.pack(fill="x", pady=(0, 18))
        
        ModernButton(mapping_controls, "➕ Add Matching Column", 
                    self.add_column_mapping, width=200, height=42,
                    gradient_colors=["#6ee7b7", "#10b981"],
                    text_color="black", bg=self.colors['section_mapping']).pack(side="left", padx=6)
        
        ModernButton(mapping_controls, "🗑️ Clear All", 
                    self.clear_mappings, width=150, height=42,
                    gradient_colors=["#fca5a5", "#ef4444"],
                    text_color="black", bg=self.colors['section_mapping']).pack(side="left", padx=6)
        
        # Scrollable mappings container
        mappings_scroll_frame = tk.Frame(self.mapping_frame, bg=self.colors['section_mapping'])
        mappings_scroll_frame.pack(fill="both", expand=True)
        
        mappings_canvas = tk.Canvas(mappings_scroll_frame, bg=self.colors['section_mapping'], 
                                   highlightthickness=0)
        mappings_scrollbar = ttk.Scrollbar(mappings_scroll_frame, orient="vertical", 
                                          command=mappings_canvas.yview)
        
        self.mappings_container = tk.Frame(mappings_canvas, bg=self.colors['section_mapping'])
        self.mappings_container.bind("<Configure>", 
                                    lambda e: mappings_canvas.configure(
                                        scrollregion=mappings_canvas.bbox("all")))
        
        mappings_canvas.create_window((0, 0), window=self.mappings_container, anchor="nw")
        mappings_canvas.configure(yscrollcommand=mappings_scrollbar.set)
        
        mappings_canvas.pack(side="left", fill="both", expand=True)
        mappings_scrollbar.pack(side="right", fill="y")
        
        self.mapping_info = tk.Label(self.mappings_container, 
                                     text="📌 Load both Excel files to start\nconfiguring column matches",
                                     fg=self.colors['text_secondary'], 
                                     bg=self.colors['section_mapping'], 
                                     font=("Segoe UI", 12), pady=50, justify="center")
        self.mapping_info.pack()
        
        # Action Button
        action_frame = tk.Frame(left_frame, bg=self.colors['bg_light'])
        action_frame.pack(fill="x", pady=(0, 12))
        
        self.process_btn_widget = ModernButton(action_frame, "Process & Merge Files", 
                                              self.process_files, width=280, height=50,
                                              gradient_colors=["#fbbf24", "#f59e0b"],
                                              text_color="black", bg=self.colors['bg_light'])
        self.process_btn_widget.pack()
        
        # === RIGHT SIDE CONTENT ===
        
        # File Columns Info with light sky blue background
        columns_frame = tk.LabelFrame(right_frame, text="  📋 File Columns  ", 
                                     font=("Segoe UI", 12, "bold"), 
                                     bg=self.colors['section_info'], 
                                     fg=self.colors['text_primary'], 
                                     padx=18, pady=18, relief="solid", bd=2,
                                     labelanchor="n")
        columns_frame.pack(fill="x", pady=(0, 18))
        
        # Primary Columns
        tk.Label(columns_frame, text="Primary File Columns:", 
                font=("Segoe UI", 10, "bold"), bg=self.colors['section_info'], 
                fg=self.colors['text_primary'], anchor="w").pack(fill="x", pady=(6, 3))
        
        self.primary_cols_text = tk.Text(columns_frame, height=4, bg="#ffffff", 
                                    fg=self.colors['text_primary'],
                                    font=("Segoe UI", 10), wrap="word", 
                                    relief="solid", bd=1, state="disabled")
        self.primary_cols_text.pack(fill="x", pady=(0, 12))
        
        # Reference Columns
        tk.Label(columns_frame, text="Reference File Columns:", 
                font=("Segoe UI", 10, "bold"), bg=self.colors['section_info'], 
                fg=self.colors['text_primary'], anchor="w").pack(fill="x", pady=(6, 3))
        
        self.ref_cols_text = tk.Text(columns_frame, height=4, bg="#ffffff", 
                                    fg=self.colors['text_primary'],
                                    font=("Segoe UI", 10), wrap="word", 
                                    relief="solid", bd=1, state="disabled")
        self.ref_cols_text.pack(fill="x")
        
        # Progress Frame with light green background
        progress_frame = tk.LabelFrame(right_frame, text="  📊 Progress & Statistics  ", 
                                      font=("Segoe UI", 12, "bold"), 
                                      bg=self.colors['section_progress'], 
                                      fg=self.colors['text_primary'], 
                                      padx=18, pady=18, relief="solid", bd=2,
                                      labelanchor="n")
        progress_frame.pack(fill="x", pady=(0, 18))
        
        # Matching Progress
        match_label = tk.Label(progress_frame, text="Matching Rows:", 
                              font=("Segoe UI", 10, "bold"), 
                              bg=self.colors['section_progress'], 
                              fg=self.colors['text_primary'], anchor="w")
        match_label.pack(fill="x", pady=(6, 3))
        
        self.match_progress = ttk.Progressbar(progress_frame, length=420, mode='determinate')
        self.match_progress.pack(fill="x", pady=6)
        
        self.match_status = tk.Label(progress_frame, text="Waiting to start...", 
                                    font=("Segoe UI", 10), 
                                    bg=self.colors['section_progress'], 
                                    fg=self.colors['text_secondary'], anchor="w")
        self.match_status.pack(fill="x", pady=(0, 12))
        
        # Merging Progress
        merge_label = tk.Label(progress_frame, text="Merging Data:", 
                              font=("Segoe UI", 10, "bold"), 
                              bg=self.colors['section_progress'], 
                              fg=self.colors['text_primary'], anchor="w")
        merge_label.pack(fill="x", pady=(6, 3))
        
        self.merge_progress = ttk.Progressbar(progress_frame, length=420, mode='determinate')
        self.merge_progress.pack(fill="x", pady=6)
        
        self.merge_status = tk.Label(progress_frame, text="Waiting to start...", 
                                    font=("Segoe UI", 10), 
                                    bg=self.colors['section_progress'], 
                                    fg=self.colors['text_secondary'], anchor="w")
        self.merge_status.pack(fill="x")
        
        # Log Frame with light pink background
        log_frame = tk.LabelFrame(right_frame, text="  📋 Activity Log  ", 
                                 font=("Segoe UI", 12, "bold"), 
                                 bg=self.colors['section_log'], 
                                 fg=self.colors['text_primary'], 
                                 padx=12, pady=12, relief="solid", bd=2,
                                 labelanchor="n")
        log_frame.pack(fill="both", expand=True)
        
        log_scroll = tk.Scrollbar(log_frame)
        log_scroll.pack(side="right", fill="y")
        
        self.log_text = tk.Text(log_frame, bg="#ffffff", 
                               fg=self.colors['text_primary'],
                               font=("Consolas", 10), wrap="word", relief="solid",
                               yscrollcommand=log_scroll.set, state="disabled", bd=1)
        self.log_text.pack(fill="both", expand=True)
        log_scroll.config(command=self.log_text.yview)
        
        self.log_message("Application started. Ready to process files.", "info")
        
        # Initial column display
        self.update_columns_display()
    
    def update_columns_display(self):
        """Update the columns display in the right panel"""
        # Primary Columns
        self.primary_cols_text.config(state="normal")
        self.primary_cols_text.delete(1.0, "end")
        if self.primary_df is not None:
            cols = ", ".join(self.primary_df.columns)
            self.primary_cols_text.insert(1.0, cols)
        else:
            self.primary_cols_text.insert(1.0, "No file loaded")
        self.primary_cols_text.config(state="disabled")
        
        # Reference Columns
        self.ref_cols_text.config(state="normal")
        self.ref_cols_text.delete(1.0, "end")
        if self.ref_df is not None:
            cols = ", ".join(self.ref_df.columns)
            self.ref_cols_text.insert(1.0, cols)
        else:
            self.ref_cols_text.insert(1.0, "No file loaded")
        self.ref_cols_text.config(state="disabled")
    
    def select_primary_file(self):
        file = filedialog.askopenfilename(
            title="Select Primary Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file:
            try:
                self.log_message(f"Loading Primary file: {Path(file).name}", "info")
                self.primary_file = file
                self.primary_df = pd.read_excel(file)
                self.primary_label.config(text=Path(file).name, fg=self.colors['success'])
                self.log_message(f"✓ Primary file loaded: {len(self.primary_df)} rows, {len(self.primary_df.columns)} columns", "success")
                self.update_columns_display()
                self.update_mapping_options()
            except Exception as e:
                self.log_message(f"✗ Error loading Primary file: {str(e)}", "error")
                messagebox.showerror("Error", f"Error loading Primary file:\n{str(e)}")
    
    def select_ref_file(self):
        file = filedialog.askopenfilename(
            title="Select Reference Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file:
            try:
                self.log_message(f"Loading Reference file: {Path(file).name}", "info")
                self.ref_file = file
                self.ref_df = pd.read_excel(file)
                self.ref_label.config(text=Path(file).name, fg=self.colors['success'])
                self.log_message(f"✓ Reference file loaded: {len(self.ref_df)} rows, {len(self.ref_df.columns)} columns", "success")
                self.update_columns_display()
                self.update_mapping_options()
            except Exception as e:
                self.log_message(f"✗ Error loading Reference file: {str(e)}", "error")
                messagebox.showerror("Error", f"Error loading Reference file:\n{str(e)}")
    
    def update_mapping_options(self):
        if self.primary_df is not None and self.ref_df is not None:
            self.mapping_info.pack_forget()
            self.log_message("Both files loaded. Ready to configure column matching.", "success")
            if not self.column_mappings:
                self.add_column_mapping()
    
    def add_column_mapping(self):
        if self.primary_df is None or self.ref_df is None:
            self.log_message("⚠ Please load both Excel files first", "warning")
            messagebox.showwarning("Warning", "Please load both Excel files first!")
            return
        
        mapping_row = tk.Frame(self.mappings_container, bg="#ffffff", 
                              relief="solid", bd=2)
        mapping_row.pack(fill="x", pady=6, padx=6)
        
        inner_frame = tk.Frame(mapping_row, bg="#ffffff")
        inner_frame.pack(fill="x", padx=18, pady=15)
        
        tk.Label(inner_frame, text="Primary Column:", font=("Segoe UI", 10, "bold"), 
                bg="#ffffff", fg=self.colors['text_primary'], 
                width=15, anchor="w").pack(side="left", padx=6)
        
        primary_combo = ttk.Combobox(inner_frame, values=list(self.primary_df.columns), 
                                 state="readonly", width=25, font=("Segoe UI", 10))
        primary_combo.pack(side="left", padx=6)
        
        tk.Label(inner_frame, text="⟷", font=("Segoe UI", 18, "bold"), 
                bg="#ffffff", fg=self.colors['primary']).pack(side="left", padx=12)
        
        tk.Label(inner_frame, text="Ref Column:", font=("Segoe UI", 10, "bold"), 
                bg="#ffffff", fg=self.colors['text_primary'], 
                width=13, anchor="w").pack(side="left", padx=6)
        
        ref_combo = ttk.Combobox(inner_frame, values=list(self.ref_df.columns), 
                                state="readonly", width=25, font=("Segoe UI", 10))
        ref_combo.pack(side="left", padx=6)
        
        remove_btn = tk.Button(inner_frame, text="✕", 
                              command=lambda: self.remove_mapping(mapping_row),
                              bg=self.colors['danger'], fg="white", 
                              font=("Segoe UI", 12, "bold"), 
                              width=3, cursor="hand2", relief="flat",
                              activebackground=self.colors['danger_hover'],
                              activeforeground="white", bd=0)
        remove_btn.pack(side="right", padx=6)
        
        self.column_mappings.append({
            'frame': mapping_row,
            'primary_combo': primary_combo,
            'ref_combo': ref_combo
        })
        
        self.log_message(f"Column mapping slot #{len(self.column_mappings)} added", "info")
    
    def remove_mapping(self, frame):
        mapping = next((m for m in self.column_mappings if m['frame'] == frame), None)
        if mapping:
            self.column_mappings.remove(mapping)
            frame.destroy()
            self.log_message(f"Column mapping removed. {len(self.column_mappings)} remaining.", "info")
        
        if not self.column_mappings and self.primary_df is not None and self.ref_df is not None:
            self.mapping_info.pack()
    
    def clear_mappings(self):
        count = len(self.column_mappings)
        for mapping in self.column_mappings:
            mapping['frame'].destroy()
        self.column_mappings = []
        
        if self.primary_df is not None and self.ref_df is not None:
            self.mapping_info.pack()
        
        self.log_message(f"All {count} column mappings cleared", "warning")
    
    def process_files(self):
        if self.primary_df is None or self.ref_df is None:
            self.log_message("✗ Cannot process: Both files must be loaded", "error")
            messagebox.showerror("Error", "Please load both Excel files first!")
            return
        
        if not self.column_mappings:
            self.log_message("✗ Cannot process: No column mappings configured", "error")
            messagebox.showerror("Error", "Please add at least one column mapping!")
            return
        
        match_pairs = []
        for i, mapping in enumerate(self.column_mappings):
            primary_col = mapping['primary_combo'].get()
            ref_col = mapping['ref_combo'].get()
            
            if not primary_col or not ref_col:
                self.log_message(f"✗ Mapping #{i+1} incomplete", "error")
                messagebox.showerror("Error", "Please select columns for all mappings!")
                return
            
            match_pairs.append((primary_col, ref_col))
            self.log_message(f"Match pair #{i+1}: '{primary_col}' ⟷ '{ref_col}'", "info")
        
        try:
            self.log_message("=" * 50, "info")
            self.log_message("Starting matching process...", "info")
            
            # Reset progress
            self.match_progress['value'] = 0
            self.merge_progress['value'] = 0
            self.match_status.config(text="Starting...", fg=self.colors['text_secondary'])
            self.merge_status.config(text="Waiting...", fg=self.colors['text_secondary'])
            
            result_df = self.primary_df.copy()
            matched_ref_cols = [pair[1] for pair in match_pairs]
            ref_additional_cols = [col for col in self.ref_df.columns if col not in matched_ref_cols]
            
            for col in ref_additional_cols:
                if col not in result_df.columns:
                    result_df[col] = ""
            
            self.log_message(f"Additional columns to merge: {len(ref_additional_cols)}", "info")
            
            # Matching phase
            matched_count = 0
            total_rows = len(self.primary_df)
            
            for idx, primary_row in self.primary_df.iterrows():
                match_condition = pd.Series([True] * len(self.ref_df))
                
                for primary_col, ref_col in match_pairs:
                    primary_val = self.normalize_value(primary_row[primary_col])
                    ref_vals = self.ref_df[ref_col].apply(self.normalize_value)
                    match_condition = match_condition & (ref_vals == primary_val)
                
                matching_indices = self.ref_df[match_condition].index
                
                if len(matching_indices) > 0:
                    matched_count += 1
                    matched_row = self.ref_df.loc[matching_indices[0]]
                    
                    for col in ref_additional_cols:
                        result_df.at[idx, col] = matched_row[col]
                
                # Update progress
                progress = ((idx + 1) / total_rows) * 100
                self.match_progress['value'] = progress
                self.match_status.config(
                    text=f"Processed {idx + 1}/{total_rows} rows | Matched: {matched_count}",
                    fg=self.colors['success']
                )
                
                if (idx + 1) % 10 == 0 or idx == total_rows - 1:
                    self.root.update()
            
            self.log_message(f"✓ Matching complete: {matched_count}/{total_rows} rows matched", "success")
            
            if matched_count == 0:
                self.log_message("⚠ WARNING: No rows matched! Check your column mappings.", "warning")
                messagebox.showwarning("No Matches Found", 
                    "No rows were matched between the files.\n\n"
                    "Please verify:\n"
                    "• Column mappings are correct\n"
                    "• Data formats match between files\n"
                    "• There are actually matching records")
            
            # Save file
            self.merge_status.config(text="Preparing to save file...", fg=self.colors['warning'])
            self.root.update()
            
            output_file = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Merged File As",
                initialfile="merged_output.xlsx"
            )
            
            if output_file:
                self.merge_status.config(text="Saving file...", fg=self.colors['warning'])
                self.merge_progress['value'] = 50
                self.root.update()
                
                result_df.to_excel(output_file, index=False)
                
                self.merge_progress['value'] = 100
                self.merge_status.config(text=f"File saved successfully!", fg=self.colors['success'])
                
                self.log_message(f"✓ File saved: {Path(output_file).name}", "success")
                self.log_message(f"Total rows: {total_rows}", "info")
                self.log_message(f"Matched rows: {matched_count}", "success")
                self.log_message(f"Unmatched rows: {total_rows - matched_count}", "warning")
                self.log_message("=" * 50, "info")
                
                match_rate = (matched_count/total_rows*100) if total_rows > 0 else 0
                success_msg = (f"✅ Files merged successfully!\n\n"
                              f"📊 Statistics:\n"
                              f"   • Total rows: {total_rows}\n"
                              f"   • Matched: {matched_count}\n"
                              f"   • Unmatched: {total_rows - matched_count}\n"
                              f"   • Match rate: {match_rate:.1f}%\n\n"
                              f"💾 Saved to:\n{output_file}")
                
                messagebox.showinfo("Success", success_msg)
            else:
                self.log_message("⚠ Save cancelled by user", "warning")
                self.merge_status.config(text="Save cancelled", fg=self.colors['text_secondary'])
        
        except Exception as e:
            self.log_message(f"✗ CRITICAL ERROR: {str(e)}", "error")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
            self.match_status.config(text="Error occurred", fg=self.colors['danger'])
            self.merge_status.config(text="Process failed", fg=self.colors['danger'])

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMatcherApp(root)
    root.mainloop()
