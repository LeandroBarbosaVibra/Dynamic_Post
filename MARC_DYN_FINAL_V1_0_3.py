#!/usr/bin/env python3
"""
Marc Dynamic Analysis Tool - Enhanced with Auto-Detection
========================================================

Vibracoustic - European FEA Department
Marc Mentat Dynamic Frequency Analysis with Professional UI and Auto-Detection

Author: Leandro Barbosa

Version: 1.0.3 
"""

import os
import glob
import re
import csv
import math
import subprocess
from collections import defaultdict

# Guideline PDF path (network location)
GUIDELINE_PDF_PATH = r"\\frafil002\VC_FEA\VC-Marc_Post\Marc_Tools_Guideline\Marc_Dynamic_Analysis_User_Guide.html"

# Try to import Marc Post API
try:
    import py_post
    MARC_POST_AVAILABLE = True
except ImportError as e:
    MARC_POST_AVAILABLE = False
    py_post = None

# Try to import Marc Mentat API
try:
    from py_mentat import py_connect, py_get_string, py_disconnect
    MARC_API_AVAILABLE = True
except ImportError as e:
    MARC_API_AVAILABLE = False
    def py_connect():
        raise Exception("Marc API not available")
    def py_get_string(cmd):
        raise Exception("Marc API not available")
    def py_disconnect():
        pass

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Global variables for file management
current_post_file = None
current_dat_file = None

_FREQ_HZ_RE = re.compile(r'([-+]?\d+(?:[.,]\d+)?)\s*hz\b', re.IGNORECASE)
_FREQ_KEY_RE = re.compile(r'(?:freq(?:uency)?)\s*[:=]?\s*([-+]?\d+(?:[.,]\d+)?)', re.IGNORECASE)
_FREQ_RAD_RE = re.compile(r'([-+]?\d+(?:[.,]\d+)?)\s*(?:rad\s*/\s*s|rad/s|rad\s*s-1|rad\s*s\^-1)', re.IGNORECASE)
_INCREMENT_ID_ACCESSORS = ("increment", "increment_id", "state_id", "current_increment")
_STATE_FREQ_ACCESSORS = (
    "frequency", "freq", "harmonic_frequency", "state_frequency",
    "increment_frequency", "current_frequency", "excitation_frequency",
    "load_frequency", "analysis_frequency"
)


def open_guideline_pdf():
    """Open the guideline PDF document"""
    try:
        if os.path.exists(GUIDELINE_PDF_PATH):
            os.startfile(GUIDELINE_PDF_PATH)
        else:
            subprocess.Popen(['explorer', GUIDELINE_PDF_PATH], shell=True)
    except Exception as e:
        messagebox.showwarning("Guideline", 
            f"Could not open guideline PDF.\n\nPath: {GUIDELINE_PDF_PATH}\n\nError: {e}")


def get_dat_from_t16(t16_file):
    """
    Automatically get the .dat file path based on .t16 file.
    The .dat file has the same name and is in the same folder as the .t16 file.
    Returns the .dat path if it exists, None otherwise.
    """
    if not t16_file:
        return None
    dat_file = os.path.splitext(t16_file)[0] + '.dat'
    if os.path.exists(dat_file):
        return dat_file
    return None


# BUTTON STYLING FUNCTIONS
def create_styled_button(parent, text, command, style="default", width=None, **kwargs):
    """Create a styled button with consistent professional appearance"""
    style_configs = {
        "default":  {'bg': '#4CAF50', 'fg': 'white', 'font': ('Arial', 9, 'bold')},
        "primary":  {'bg': '#2196F3', 'fg': 'white', 'font': ('Arial', 10, 'bold')},
        "danger":   {'bg': '#d4542a', 'fg': 'white', 'font': ('Arial', 9, 'bold')},
        "warning":  {'bg': '#FF9800', 'fg': 'white', 'font': ('Arial', 9, 'bold')},
        "secondary":{'bg': '#757575', 'fg': 'white', 'font': ('Arial', 9)},
        "success":  {'bg': '#4CAF50', 'fg': 'white', 'font': ('Arial', 9, 'bold')}
    }
    config = style_configs.get(style, style_configs["default"])
    config.update({
        'relief': 'raised',
        'bd': 3,
        'cursor': 'hand2',
        'activebackground': _darken_color(config['bg']),
        'activeforeground': 'white'
    })
    if width:
        config['width'] = width
    config.update(kwargs)
    return tk.Button(parent, text=text, command=command, **config)

def _darken_color(hex_color):
    """Darken a hex color for active state"""
    try:
        hex_color = hex_color.lstrip('#')
        rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        darker_rgb = tuple(max(0, int(c * 0.8)) for c in rgb)
        return f"#{darker_rgb[0]:02x}{darker_rgb[1]:02x}{darker_rgb[2]:02x}"
    except:
        return "#333333"


def _scalar_to_float(value):
    """Convert Marc API values to a finite float when possible."""
    if value is None:
        return None
    try:
        if isinstance(value, (list, tuple)):
            if not value:
                return None
            value = value[0]
        elif hasattr(value, "shape"):
            arr = value.ravel()
            if arr.size == 0:
                return None
            value = arr[0]
        elif hasattr(value, "__len__") and not isinstance(value, (str, bytes)):
            seq = list(value)
            if not seq:
                return None
            value = seq[0]
    except Exception:
        pass

    try:
        out = float(value)
        if math.isfinite(out):
            return out
    except Exception:
        pass
    return None


def _call_with_fallback(func, increment_index):
    """Try Marc API accessors with the common call signatures."""
    for args in ((), (increment_index,), (max(0, int(increment_index) - 1),)):
        try:
            return func(*args)
        except TypeError:
            continue
        except Exception:
            return None
    return None


def _read_post_numeric(post_file, accessor, increment_index):
    """Read a numeric value from the Marc post object."""
    if not accessor or not hasattr(post_file, accessor):
        return None
    try:
        obj = getattr(post_file, accessor)
    except Exception:
        return None

    raw = _call_with_fallback(obj, increment_index) if callable(obj) else obj
    return _scalar_to_float(raw)


def _extract_frequency_from_title(title):
    """Parse frequency information from the increment title."""
    txt = str(title or "")
    if not txt:
        return None

    match = _FREQ_RAD_RE.search(txt)
    if match:
        try:
            return float(match.group(1).replace(",", ".")) / (2.0 * math.pi)
        except Exception:
            pass

    match = _FREQ_HZ_RE.search(txt)
    if match:
        try:
            return float(match.group(1).replace(",", "."))
        except Exception:
            pass

    match = _FREQ_KEY_RE.search(txt)
    if match:
        try:
            return float(match.group(1).replace(",", "."))
        except Exception:
            pass

    return None


def resolve_post_increment_id(post_file, post_index, default_value=None):
    """Return the real increment ID exposed by Marc for a given internal state index."""
    fallback = post_index if default_value is None else default_value

    for accessor in _INCREMENT_ID_ACCESSORS:
        if not hasattr(post_file, accessor):
            continue
        try:
            obj = getattr(post_file, accessor)
        except Exception:
            continue

        raw = _call_with_fallback(obj, post_index) if callable(obj) else obj
        if raw is None:
            continue

        value = _scalar_to_float(raw)
        if value is None:
            try:
                value = float(str(raw).strip())
            except Exception:
                continue

        try:
            return int(round(value))
        except Exception:
            continue

    return fallback


def resolve_increment_frequency_hz(post_file, increment_index, state_title=""):
    """Resolve the excitation frequency in Hz for a post state."""
    title_freq = _extract_frequency_from_title(state_title)
    if title_freq is not None:
        return title_freq

    for accessor in _STATE_FREQ_ACCESSORS:
        freq = _read_post_numeric(post_file, accessor, increment_index)
        if freq is not None and abs(freq) > 1e-12:
            return freq

    return None


class MarcMentatSession:
    """Manages Marc Mentat session and file detection"""
    def __init__(self):
        self.connected = False
        self.active_file = None
        self.dat_file = None
        self.working_directory = None
    
    def __enter__(self):
        if not MARC_API_AVAILABLE:
            return None
        try:
            py_connect()
            self.connected = True
            self.working_directory = self.detect_working_directory()
            self.active_file = self.detect_active_t16_file()
            if self.active_file and not isinstance(self.active_file, list):
                # Auto-detect .dat file based on .t16 file
                self.dat_file = get_dat_from_t16(self.active_file)
            return self
        except Exception as e:
            return None
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.connected:
            try:
                py_disconnect()
            except:
                pass
    
    def detect_working_directory(self):
        try:
            working_dir = py_get_string('getcwd()')
            if working_dir and working_dir.strip():
                return working_dir.strip()
        except Exception as e:
            pass
        working_dir = os.getcwd()
        return working_dir
    
    def detect_active_t16_file(self):
        try:
            fname = py_get_string('filename()')
            if fname and fname.strip():
                if not fname.endswith('.t16'):
                    fname = fname + '.t16'
                full_path = os.path.join(self.working_directory, fname) if self.working_directory else fname
                if os.path.exists(full_path):
                    return full_path
        except Exception as e:
            pass
        search_dir = self.working_directory or os.getcwd()
        return self.select_t16_file_from_directory(search_dir)
    
    def select_t16_file_from_directory(self, directory):
        try:
            t16_files = glob.glob(os.path.join(directory, "*.t16"))
            if len(t16_files) == 0:
                return None
            elif len(t16_files) == 1:
                return t16_files[0]
            else:
                return t16_files
        except Exception:
            return None


class FileSelectionWindow:
    """Window for selecting .t16 file when multiple exist (dat is auto-detected)"""
    def __init__(self, parent, t16_files, callback):
        self.window = tk.Toplevel(parent)
        self.window.title("Select File - Marc Dynamic Analysis")
        self.window.geometry("825x500")
        self.window.grab_set()
        self.callback = callback
        self.t16_files = t16_files if isinstance(t16_files, list) else [t16_files] if t16_files else []
        self.selected_t16 = None
        self.setup_ui()
    
    def setup_ui(self):
        title_frame = tk.Frame(self.window, bg='#d4542a', relief='raised', bd=3)
        title_frame.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(title_frame, text="Marc Dynamic Analysis - File Selection", font=('Arial', 14, 'bold'),
                 fg='white', bg='#d4542a', pady=8).pack()

        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main_frame, text="Multiple .t16 files found - Please select:", font=("Arial", 12, "bold")).pack(pady=10)
        
        selection_frame = ttk.LabelFrame(main_frame, text="Current Selection", padding="10")
        selection_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(selection_frame, text="Selected T16 file:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W, padx=5)
        self.selected_t16_label = ttk.Label(selection_frame, text="None selected", foreground="gray")
        self.selected_t16_label.grid(row=0, column=1, sticky=tk.W, padx=10)
        ttk.Label(selection_frame, text="DAT file:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky=tk.W, padx=5)
        self.selected_dat_label = ttk.Label(selection_frame, text="(auto-detected)", foreground="gray", font=("Arial", 9, "italic"))
        self.selected_dat_label.grid(row=1, column=1, sticky=tk.W, padx=10)
        
        if len(self.t16_files) > 1:
            t16_frame = ttk.LabelFrame(main_frame, text="Select .t16 File", padding="10")
            t16_frame.pack(fill=tk.BOTH, expand=True, pady=5)
            self.t16_listbox = tk.Listbox(t16_frame, height=8)
            t16_scrollbar = ttk.Scrollbar(t16_frame, orient=tk.VERTICAL, command=self.t16_listbox.yview)
            self.t16_listbox.configure(yscrollcommand=t16_scrollbar.set)
            for t16_file in self.t16_files:
                self.t16_listbox.insert(tk.END, os.path.basename(t16_file))
            self.t16_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            t16_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.t16_listbox.selection_set(0)
            self.t16_listbox.bind('<<ListboxSelect>>', self.on_t16_selection_change)
            self.selected_t16 = self.t16_files[0]
            self.selected_t16_label.config(text=os.path.basename(self.selected_t16), foreground="darkgreen")
            self.update_dat_label()
        
        info_label = ttk.Label(main_frame, text="Note: The .dat file is automatically detected (same name as .t16)", 
                               font=("Arial", 9, "italic"), foreground="gray")
        info_label.pack(pady=5)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        create_styled_button(button_frame, "OK", self.confirm_selection, style="primary", width=12).pack(side=tk.RIGHT, padx=5)
        create_styled_button(button_frame, "Browse...", self.browse_files, style="secondary", width=12).pack(side=tk.RIGHT, padx=5)
        create_styled_button(button_frame, "Cancel", self.window.destroy, style="secondary", width=12).pack(side=tk.RIGHT)
    
    def update_dat_label(self):
        if self.selected_t16:
            dat_file = get_dat_from_t16(self.selected_t16)
            if dat_file:
                self.selected_dat_label.config(text=os.path.basename(dat_file), foreground="darkgreen", font=("Arial", 10))
            else:
                self.selected_dat_label.config(text="Not found (will be auto-detected)", foreground="orange", font=("Arial", 9, "italic"))
    
    def on_t16_selection_change(self, event):
        selection = self.t16_listbox.curselection()
        if selection:
            self.selected_t16 = self.t16_files[selection[0]]
            self.selected_t16_label.config(text=os.path.basename(self.selected_t16), foreground="darkgreen")
            self.update_dat_label()
    
    def confirm_selection(self):
        if hasattr(self, 't16_listbox'):
            t16_selection = self.t16_listbox.curselection()
            if t16_selection:
                self.selected_t16 = self.t16_files[t16_selection[0]]
        elif len(self.t16_files) == 1:
            self.selected_t16 = self.t16_files[0]
        if not self.selected_t16:
            messagebox.showerror("Error", "Please select a .t16 file")
            return
        self.callback(self.selected_t16)
        self.window.destroy()
    
    def browse_files(self):
        t16_file = filedialog.askopenfilename(title="Select .t16 file", filetypes=[("T16 files", "*.t16"), ("All files", "*.*")])
        if t16_file:
            self.callback(t16_file)
            self.window.destroy()


class MainApplication:
    """Main application with auto-detection and professional interface"""
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Vibracoustic - Marc Dynamic Analysis Tool")
        self.root.geometry("725x785")
        self.root.configure(bg='#f0f0f0')
        self.session = None
        self.increment_data = None  # Cache for increment data
        self.setup_ui()
        try:
            self.root.update_idletasks()
            self.root.update()
        except tk.TclError as e:
            if "ThemeChanged" not in str(e) and "application has been destroyed" not in str(e):
                raise
        self.root.after(500, self.initialize_session_safe)

    def setup_ui(self):
        # Title Frame with Guideline Button
        title_frame = tk.Frame(self.root, bg='#d4542a', relief='raised', bd=3)
        title_frame.pack(fill=tk.X, padx=10, pady=10)
        
        title_content = tk.Frame(title_frame, bg='#d4542a')
        title_content.pack(fill=tk.X, padx=(5, 90), pady=5)
        
        guideline_btn = tk.Button(title_frame, text="Guideline", command=open_guideline_pdf,
                                  font=('Arial', 8, 'bold'), bg='#FFEB3B', fg='#333333',
                                  relief='raised', bd=2, cursor='hand2', padx=8, pady=2,
                                  activebackground='#FFC107', activeforeground='#333333')
        guideline_btn.place(relx=1.0, x=-8, y=8, anchor='ne')
        
        title_center = tk.Frame(title_content, bg='#d4542a')
        title_center.pack(fill=tk.X, expand=True)
        
        tk.Label(title_center, text="Marc Dynamic Analysis Tool", font=('Arial', 16, 'bold'),
                 fg='white', bg='#d4542a', pady=8).pack(anchor='center')
        tk.Label(title_center, text="Vibracoustic - European FEA Department", font=('Arial', 11, 'bold'),
                 fg='white', bg='#d4542a', pady=5).pack(anchor='center')
        
        file_frame = ttk.LabelFrame(self.root, text="File Information", padding="10")
        file_frame.pack(fill=tk.X, padx=10, pady=10)
        info_grid = ttk.Frame(file_frame)
        info_grid.pack(fill=tk.X)
        ttk.Label(info_grid, text="T16 File:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.t16_file_label = ttk.Label(info_grid, text="Not loaded", foreground="gray")
        self.t16_file_label.grid(row=0, column=1, sticky=tk.W)
        ttk.Label(info_grid, text="Status:").grid(row=0, column=2, sticky=tk.W, padx=(20, 10))
        self.status_label = ttk.Label(info_grid, text="Ready", foreground="blue")
        self.status_label.grid(row=0, column=3, sticky=tk.W)
        ttk.Label(info_grid, text="DAT File:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
        self.dat_file_label = ttk.Label(info_grid, text="Not loaded", foreground="gray")
        self.dat_file_label.grid(row=1, column=1, sticky=tk.W)
        
        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill=tk.X, padx=10, pady=20)
        create_styled_button(control_frame, "Load Files", self.load_files, style="primary", width=12).pack(side=tk.LEFT, padx=5)
        create_styled_button(control_frame, "Select Files Manually", self.select_files_manually, style="secondary", width=18).pack(side=tk.LEFT, padx=5)
        create_styled_button(control_frame, "Start Dynamic Analysis", self.start_analysis, style="success", width=20).pack(side=tk.LEFT, padx=5)
        create_styled_button(control_frame, "Cancel", self.safe_exit, style="danger", width=12).pack(side=tk.LEFT, padx=5)
        
        info_frame = ttk.LabelFrame(self.root, text="Marc Dynamic Analysis Information", padding="15")
        info_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        info_text = """Marc Dynamic Analysis Tool - V1.0.3

This tool performs frequency response analysis on Marc Mentat solution files:

* Automatically detects active .t16 and .dat files from Marc Mentat
* The .dat file is auto-detected (same name as .t16, same folder)
* Shows only harmonic increments in the selection list
* Displays the real increment IDs read from Marc/Mentat
* Extracts frequency-dependent displacement and reaction force data  
* Supports node sets and individual node selection
* Processes complete frequency sweeps with phase information
* Exports results to CSV format for post-processing

Usage:
1. Ensure Marc Mentat is running with a loaded model (optional)
2. Click 'Load Files' to auto-detect files or use manual selection
3. Click 'Start Dynamic Analysis' to configure and run analysis

Author: Leandro Barbosa"""
        tk.Label(info_frame, text=info_text, font=('Arial', 9), fg='#333333', bg='#f5f5f5',
                 justify=tk.LEFT, wraplength=650).pack(anchor='w')
        
        # Progress Frame (ALWAYS VISIBLE)
        self.progress_frame = ttk.LabelFrame(self.root, text="Loading Progress", padding="10")
        self.progress_frame.pack(fill=tk.X, padx=10, pady=5)
        
        progress_label_frame = tk.Frame(self.progress_frame)
        progress_label_frame.pack(fill=tk.X, anchor=tk.W)
        self.progress_label = tk.Label(progress_label_frame, text="Ready", fg="gray", 
                                       font=('Arial', 9), anchor='w', width=80)
        self.progress_label.pack(anchor=tk.W)
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate', length=400)
        self.progress_bar.pack(fill=tk.X, pady=5)
        self.progress_bar['value'] = 0
        
        detail_label_frame = tk.Frame(self.progress_frame)
        detail_label_frame.pack(fill=tk.X, anchor=tk.W)
        self.progress_detail = tk.Label(detail_label_frame, text="", fg="gray",
                                        font=('Arial', 9), anchor='w', width=80)
        self.progress_detail.pack(anchor=tk.W)
        
        self.status_var = tk.StringVar(value="Ready - Marc Dynamic Analysis Tool")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.root.protocol("WM_DELETE_WINDOW", self.safe_exit)

    # =========================================================================
    # PROGRESS BAR METHODS
    # =========================================================================
    def set_progress_label(self, text, color="blue"):
        self.progress_label.configure(text=text, fg=color)
        self.root.update_idletasks()
    
    def set_progress_detail(self, text):
        self.progress_detail.configure(text=text)
        self.root.update_idletasks()

    def reset_progress(self):
        self.progress_bar['value'] = 0
        self.set_progress_label("Ready", "gray")
        self.set_progress_detail("")
        self.root.update_idletasks()

    # =========================================================================
    # HARMONIC INCREMENT LOADING WITH REAL MARC IDS
    # =========================================================================
    def load_increments_fast(self, t16_file):
        """
        Load only harmonic increments and keep the real Marc IDs.
        
        Returns a dictionary with:
        - 'indices': list of internal indices (0-based position in .t16)
        - 'ids': list of actual increment IDs exposed by Marc
        - 'labels': list of UI labels based on real IDs
        - 'label_to_index': dict mapping label -> internal index
        """
        try:
            global py_post
            if not MARC_POST_AVAILABLE or py_post is None:
                try:
                    import py_post as pp
                    py_post = pp
                except:
                    return None
            
            self.reset_progress()
            
            self.set_progress_label("Opening post file...", "blue")
            self.set_progress_detail("")
            self.progress_bar['value'] = 10
            self.root.update_idletasks()
            
            post = py_post.post_open(t16_file)
            total_inc = post.increments()
            
            if total_inc == 0:
                post.close()
                self.reset_progress()
                return {'indices': [], 'ids': [], 'labels': [], 'label_to_index': {}, 'total': 0}
            
            # Start from index 1 if there are multiple increments (skip index 0)
            start_index = 1 if total_inc > 1 else 0
            actual_count = total_inc - start_index
            
            if actual_count <= 0:
                post.close()
                self.reset_progress()
                return {'indices': [], 'ids': [], 'labels': [], 'label_to_index': {}, 'total': total_inc}
            
            self.set_progress_label("Reading increment data from solution file...", "blue")
            self.progress_bar['value'] = 25
            self.root.update_idletasks()

            increment_records = []
            progress_step = max(1, actual_count // 20)
            for offset, idx in enumerate(range(start_index, total_inc), start=1):
                post.moveto(idx)
                increment_id = resolve_post_increment_id(post, idx, idx)

                try:
                    raw_title = post.title()
                    title = str(raw_title).strip() if raw_title is not None else ""
                except Exception:
                    title = ""

                frequency_hz = resolve_increment_frequency_hz(post, idx, title)
                increment_records.append({
                    'index': idx,
                    'id': increment_id,
                    'frequency': frequency_hz
                })

                if offset == 1 or offset == actual_count or offset % progress_step == 0:
                    progress = 25 + int((offset / actual_count) * 55)
                    self.progress_bar['value'] = progress
                    self.set_progress_detail(f"Reading increment {offset}/{actual_count}...")
                    self.root.update_idletasks()

            post.close()

            self.progress_bar['value'] = 85
            self.root.update_idletasks()

            harmonic_records = [rec for rec in increment_records if rec['frequency'] is not None]
            harmonic_detection = "frequency"

            if not harmonic_records:
                total_per_id = defaultdict(int)
                for rec in increment_records:
                    total_per_id[rec['id']] += 1

                seen_per_id = defaultdict(int)
                harmonic_records = []
                for rec in increment_records:
                    seen_per_id[rec['id']] += 1
                    if total_per_id[rec['id']] > 1 and seen_per_id[rec['id']] > 1:
                        harmonic_records.append(rec)

                if harmonic_records:
                    harmonic_detection = "repeated_id"

            has_harmonics = len(harmonic_records) > 0
            increment_indices = []
            increment_ids = []
            increment_labels = []
            increment_frequencies = []
            label_to_index = {}

            if has_harmonics:
                total_per_id = defaultdict(int)
                for rec in harmonic_records:
                    total_per_id[rec['id']] += 1

                seen_per_id = defaultdict(int)
                for rec in harmonic_records:
                    increment_indices.append(rec['index'])
                    increment_ids.append(rec['id'])
                    increment_frequencies.append(rec['frequency'])

                    seen_per_id[rec['id']] += 1
                    label = str(rec['id'])
                    if total_per_id[rec['id']] > 1:
                        if rec['frequency'] is not None:
                            label = f"{rec['id']} @ {rec['frequency']:.6g} Hz"
                        else:
                            label = f"{rec['id']} (state {seen_per_id[rec['id']]})"

                    increment_labels.append(label)
                    label_to_index[label] = rec['index']

                self.set_progress_label(f"Loaded {len(harmonic_records)} harmonic increments", "green")
                if harmonic_detection == "frequency":
                    self.set_progress_detail("Only harmonic increments with real Marc IDs are shown")
                else:
                    self.set_progress_detail("Harmonic increments detected from repeated real Marc IDs")
            else:
                self.set_progress_label("No harmonic increments detected", "orange")
                self.set_progress_detail("The file does not expose harmonic increments with frequency metadata")
            
            self.progress_bar['value'] = 100
            self.root.update_idletasks()
            
            return {
                'indices': increment_indices,
                'ids': increment_ids,
                'frequencies': increment_frequencies,
                'labels': increment_labels,
                'label_to_index': label_to_index,
                'total': total_inc,
                'has_harmonics': has_harmonics,
                'harmonic_increment': increment_ids[0] if has_harmonics else None
            }
            
        except Exception as e:
            self.reset_progress()
            return None

    def initialize_session_safe(self):
        try:
            self.status_var.set("Attempting to detect Marc Mentat files...")
            self.root.update()
            self.initialize_session()
        except Exception as e:
            self.status_var.set("Ready - Use 'Load Files' or 'Select Files Manually'")

    def initialize_session(self):
        global current_post_file, current_dat_file
        try:
            self.session = MarcMentatSession()
            with self.session as session:
                if session and session.active_file:
                    if isinstance(session.active_file, list):
                        t16_files = session.active_file
                        self.status_var.set("Multiple files found - please select")
                        self.root.after(200, lambda: FileSelectionWindow(self.root, t16_files, self.on_files_selected))
                    else:
                        current_post_file = session.active_file
                        current_dat_file = get_dat_from_t16(current_post_file)
                        self.update_file_display()
                        self.status_var.set("Files automatically detected and loaded")
                else:
                    self.status_var.set("No active files detected - use Load Files or Select Manually")
        except Exception as e:
            self.status_var.set("Marc Mentat not connected - use Load Files or Select Manually")

    def load_files(self):
        global current_post_file, current_dat_file
        try:
            if current_post_file and os.path.exists(current_post_file):
                self.status_var.set("Files already loaded. Use 'Select Files Manually' to change.")
                self.update_file_display()
                return
            working_dir = os.getcwd()
            t16_files = glob.glob(os.path.join(working_dir, "*.t16"))
            if len(t16_files) == 0:
                messagebox.showinfo("No Files Found", "No .t16 files found in current directory.\nPlease use 'Select Files Manually'.", parent=self.root)
                self.status_var.set("No .t16 files found - use manual selection")
                return
            elif len(t16_files) > 1:
                self.status_var.set("Multiple files found - please select")
                FileSelectionWindow(self.root, t16_files, self.on_files_selected)
                return
            else:
                current_post_file = t16_files[0]
                current_dat_file = get_dat_from_t16(current_post_file)
                self.status_var.set("Single .t16 file found and loaded")
            
            self.update_file_display()
            self.increment_data = None
            if current_post_file:
                self.status_var.set("Files loaded successfully - Ready for dynamic analysis")
            else:
                self.status_var.set("Failed to load files")
        except Exception as e:
            error_msg = f"Failed to load files: {e}"
            messagebox.showerror("Error", error_msg, parent=self.root)
            self.status_var.set(error_msg)
    
    def update_file_display(self):
        global current_post_file, current_dat_file
        if current_post_file and os.path.exists(current_post_file):
            self.t16_file_label.config(text=os.path.basename(current_post_file), foreground="darkgreen")
            self.status_label.config(text="T16 Loaded", foreground="darkgreen")
        else:
            self.t16_file_label.config(text="Not loaded", foreground="gray")
            self.status_label.config(text="Ready", foreground="blue")
        if current_dat_file and os.path.exists(current_dat_file):
            self.dat_file_label.config(text=os.path.basename(current_dat_file), foreground="darkgreen")
        else:
            self.dat_file_label.config(text="Not found (auto-detect failed)", foreground="orange")

    def on_files_selected(self, selected_t16):
        global current_post_file, current_dat_file
        current_post_file = selected_t16
        current_dat_file = get_dat_from_t16(selected_t16)
        self.increment_data = None
        self.update_file_display()
        self.reset_progress()
        if current_dat_file:
            self.status_var.set("Files selected and loaded - Ready for analysis")
        else:
            self.status_var.set("T16 loaded, DAT not found - Node sets may not be available")

    def select_files_manually(self):
        global current_post_file, current_dat_file
        try:
            t16_file = filedialog.askopenfilename(title="Select .t16 file",
                                                  filetypes=[("T16 files", "*.t16"), ("All files", "*.*")],
                                                  parent=self.root)
            if not t16_file:
                self.status_var.set("File selection cancelled")
                return
            current_post_file = t16_file
            current_dat_file = get_dat_from_t16(t16_file)
            
            if not current_dat_file:
                messagebox.showwarning("Warning", 
                    f"DAT file not found.\n\n"
                    f"Expected: {os.path.splitext(t16_file)[0]}.dat\n\n"
                    f"Node set features may not be available.", 
                    parent=self.root)
            
            self.increment_data = None
            self.update_file_display()
            self.reset_progress()
            self.status_var.set("Files selected manually - Ready for analysis")
        except Exception as e:
            error_msg = f"Failed to select files: {e}"
            messagebox.showerror("Error", error_msg, parent=self.root)
            self.status_var.set(error_msg)

    def safe_exit(self):
        try:
            result = messagebox.askyesno("Confirm Exit",
                                         "Are you sure you want to close the Marc Dynamic Analysis Tool?\n\n"
                                         "This will NOT close Marc Mentat.",
                                         parent=self.root)
            if result:
                try:
                    self.root.quit()
                    self.root.destroy()
                except:
                    pass
        except Exception:
            try:
                self.root.destroy()
            except:
                pass

    def start_analysis(self):
        global current_post_file, current_dat_file
        if not current_post_file:
            messagebox.showerror("Error", "Please load a .t16 file first")
            return
        if not current_dat_file or not os.path.exists(current_dat_file):
            messagebox.showwarning("Warning", "DAT file not found. Node set features may not be available.")
        
        self.status_var.set("Loading increment data...")
        self.root.update()
        
        if self.increment_data is None:
            self.increment_data = self.load_increments_fast(current_post_file)
        
        if self.increment_data is None or len(self.increment_data.get('labels', [])) == 0:
            no_harmonics = self.increment_data is not None and not self.increment_data.get('has_harmonics', False)
            error_text = (
                "No harmonic increments were detected in the .t16 file.\n"
                "This workflow now lists only harmonic increments with their real Marc IDs."
                if no_harmonics else
                "Could not load increment data from the .t16 file.\n"
                "Please verify the file is valid."
            )
            messagebox.showerror("Error", 
                error_text, 
                parent=self.root)
            self.status_var.set("Failed to load increments")
            self.reset_progress()
            return
        
        self.status_var.set("Starting dynamic analysis configuration...")
        self.root.withdraw()
        try:
            self.run_dynamic_analysis()
        except Exception as e:
            messagebox.showerror("Analysis Error", f"Analysis failed: {e}")
        finally:
            self.root.deiconify()
            self.status_var.set("Analysis workflow completed")

    def run_dynamic_analysis(self):
        global current_post_file, current_dat_file
        if not MARC_POST_AVAILABLE or py_post is None:
            try:
                import py_post as pp
                globals()['py_post'] = pp
            except:
                messagebox.showerror("Error", 
                    "Marc Post API (py_post) is not available.\n\n"
                    "Please ensure:\n"
                    "1. Marc Mentat is properly installed\n"
                    "2. Python path includes Marc's Python libraries\n"
                    "3. You are running from Marc Mentat's Python environment",
                    parent=self.root)
                return
        
        os.chdir(os.path.dirname(current_post_file))
        
        params = get_parameters(self.root, self.increment_data)
        if not params:
            return

        inc = params['inc']  # This is now the internal index
        single_in = params['disp']
        rf_input = params['rf_set']
        direction = params['direction']

        progress_win, status_text = show_progress_window(self.root)
        update_progress(status_text, "=== Marc Dynamic Analysis Started ===")
        update_progress(status_text, f"Solution file: {os.path.basename(current_post_file)}")
        update_progress(status_text, f"Direction: {direction}")
        update_progress(status_text, "")

        try:
            update_progress(status_text, "Loading node definitions...")
            try:
                try:
                    single_node_id = int(single_in)
                    update_progress(status_text, f"[OK] Using node ID {single_node_id} for displacement")
                except ValueError:
                    if current_dat_file and os.path.exists(current_dat_file):
                        ids = get_nodes_from_set(current_dat_file, single_in)
                        if len(ids) != 1:
                            raise ValueError("Excitation node set must contain exactly one node.")
                        single_node_id = ids[0]
                        update_progress(status_text, f"[OK] Node {single_node_id} selected from set '{single_in}'")
                    else:
                        raise ValueError(f"Cannot resolve node set '{single_in}' - DAT file not available")

                try:
                    rf_id = int(rf_input)
                    node_ids = {rf_id}
                    update_progress(status_text, f"[OK] Using node ID {rf_id} for reaction forces")
                except ValueError:
                    if current_dat_file and os.path.exists(current_dat_file):
                        ids = get_nodes_from_set(current_dat_file, rf_input)
                        if not ids:
                            raise ValueError(f"No nodes found in set '{rf_input}'.")
                        node_ids = set(ids)
                        update_progress(status_text, f"[OK] {len(node_ids)} nodes loaded from set '{rf_input}'")
                    else:
                        raise ValueError(f"Cannot resolve node set '{rf_input}' - DAT file not available")
                        
            except Exception as e:
                update_progress(status_text, f"[ERROR] Error reading node definitions: {e}")
                messagebox.showerror("Error", f"Error reading node definitions:\n{e}")
                progress_win.destroy()
                return

            update_progress(status_text, "")
            update_progress(status_text, "Opening solution file...")
            post = py_post.post_open(current_post_file)
            total_inc = post.increments()
            
            if not (0 <= inc < total_inc):
                error_msg = f"Increment {inc} out of range (0..{total_inc-1})"
                update_progress(status_text, f"[ERROR] {error_msg}")
                messagebox.showerror("Error", error_msg)
                post.close()
                progress_win.destroy()
                return
                
            update_progress(status_text, f"[OK] Solution file opened ({total_inc} increments)")

            update_progress(status_text, "")
            update_progress(status_text, "Reading frequency data...")
            freqs = []
            for idx in range(inc, total_inc):
                post.moveto(idx)
                freqs.append(post.frequency)

            update_progress(status_text, f"[OK] Found {len(freqs)} frequency points")
            update_progress(status_text, "")
            update_progress(status_text, "=== Available Frequencies ===")
            for i, f in enumerate(freqs[:10]):
                update_progress(status_text, f"{i:3d}: {f:.6g} Hz")
            if len(freqs) > 10:
                update_progress(status_text, f"... and {len(freqs)-10} more")
            update_progress(status_text, "")

            start_idx = next((i for i, f in enumerate(freqs) if f > 0), 0)
            post.moveto(inc + start_idx)
            update_progress(status_text, f"[OK] Starting analysis from frequency {freqs[start_idx]:.6g} Hz")

            update_progress(status_text, "")
            update_progress(status_text, "Locating data fields...")
            labels = [post.node_scalar_label(i) for i in range(post.node_scalars())]
            try:
                i_mag = labels.index(f"Reaction Force {direction} Magnitude")
                i_phase = labels.index(f"Reaction Force {direction} Phase")
                i_disp = labels.index(f"Displacement {direction} Magnitude")
                update_progress(status_text, f"[OK] Located {direction}-direction data fields")
            except ValueError as e:
                error_msg = f"Required data fields not found in solution: {e}"
                update_progress(status_text, f"[ERROR] {error_msg}")
                messagebox.showerror("Error", error_msg)
                post.close()
                progress_win.destroy()
                return

            out_csv = get_unique_filename("Marc_Dyn_Solution")
            update_progress(status_text, f"[OK] Output file: {out_csv}")
            update_progress(status_text, "")
            update_progress(status_text, "Processing frequency sweep...")

            with open(out_csv, "w", newline="") as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow([
                    "Frequency(Hz)",
                    f"Displacement {direction} Magnitude",
                    "Magnitude", "RealSum", "ImagSum", "Phase(deg)"
                ])
                
                total_freqs = len(freqs) - start_idx
                for j in range(start_idx, len(freqs)):
                    idx = inc + j
                    post.moveto(idx)
                    real_sum = imag_sum = 0.0
                    
                    for n in range(post.nodes()):
                        node = post.node(n)
                        if node.id in node_ids:
                            mag = post.node_scalar(n, i_mag)
                            ph = math.radians(post.node_scalar(n, i_phase))
                            real_sum += mag * math.cos(ph)
                            imag_sum += mag * math.sin(ph)
                            
                    vec_mag = math.hypot(real_sum, imag_sum)
                    vec_ph = math.degrees(math.atan2(imag_sum, real_sum))
                    vec_ph = (vec_ph + 180) % 360 - 180

                    disp = float('nan')
                    for n in range(post.nodes()):
                        if post.node(n).id == single_node_id:
                            disp = post.node_scalar(n, i_disp)
                            break

                    writer.writerow([
                        freqs[j], disp,
                        vec_mag, real_sum, imag_sum, vec_ph
                    ])

                    progress = ((j - start_idx + 1) / total_freqs) * 100
                    if (j - start_idx + 1) % max(1, total_freqs // 10) == 0:
                        update_progress(status_text, f"  Progress: {progress:.0f}% ({freqs[j]:.3g} Hz)")

            update_progress(status_text, "")
            update_progress(status_text, "=== Analysis Complete ===")
            update_progress(status_text, f"[OK] Results saved to: {out_csv}")
            update_progress(status_text, f"[OK] Processed {total_freqs} frequency points")
            update_progress(status_text, f"[OK] Direction: {direction}")
            update_progress(status_text, "")
            update_progress(status_text, "Analysis finished successfully!")

            post.close()
            progress_win.destroy()
            
            show_completion_dialog(self.root, out_csv, total_freqs, direction)

        except Exception as e:
            error_msg = f"Analysis failed: {e}"
            update_progress(status_text, f"[ERROR] {error_msg}")
            messagebox.showerror("Analysis Error", error_msg)
            progress_win.destroy()

    def run(self):
        try:
            self.root.eval('tk::PlaceWindow . center')
            self.root.mainloop()
        except KeyboardInterrupt:
            self.safe_exit()
        except Exception as e:
            try:
                self.root.destroy()
            except:
                pass


# === HELPER FUNCTIONS ===

def get_nodes_from_set(dat_path, set_name):
    with open(dat_path, "r") as f:
        lines = f.readlines()
    start = next((i for i, line in enumerate(lines) if set_name in line), None)
    if start is None:
        raise ValueError(f"Set '{set_name}' not found in {dat_path}")
    ids = []
    for line in lines[start+1:]:
        if re.match(r"^\s*\d+", line):
            ids.extend(int(n) for n in re.findall(r"\d+", line))
        else:
            break
    return list(set(ids))

def get_node_sets_from_dat(dat_file):
    if not dat_file or not os.path.exists(dat_file):
        return []
    try:
        with open(dat_file, 'r') as f:
            content = f.read()
        node_sets = []
        for match in re.finditer(r'^\s*define\s+node\s+set\s+(\S+)', content, re.MULTILINE | re.IGNORECASE):
            set_name = match.group(1).strip()
            if set_name not in node_sets:
                node_sets.append(set_name)
        return sorted(node_sets)
    except:
        return []

def get_unique_filename(base_name):
    for i in range(1, 1000):
        filename = f"{base_name}_{i:02d}.csv"
        if not os.path.exists(filename):
            return filename
    raise RuntimeError("Too many output files, clean up the folder.")

def show_progress_window(root):
    progress_win = tk.Toplevel(root)
    progress_win.title("Marc Dynamic Analysis - Processing")
    progress_win.geometry("500x600")
    progress_win.configure(bg='#f0f0f0')
    progress_win.grab_set()
    progress_win.protocol("WM_DELETE_WINDOW", lambda: None)
    
    title_frame = tk.Frame(progress_win, bg='#2196F3', relief='raised', bd=3)
    title_frame.pack(fill=tk.X, padx=10, pady=10)
    tk.Label(title_frame, text="Marc Dynamic Analysis", font=('Arial', 14, 'bold'), 
             fg='white', bg='#2196F3', pady=8).pack()
    tk.Label(title_frame, text="Processing Frequency Sweep", font=('Arial', 10), 
             fg='white', bg='#2196F3', pady=5).pack()
    
    content_frame = tk.Frame(progress_win, bg='#f0f0f0')
    content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
    
    status_text = tk.Text(content_frame, height=10, width=60, font=('Consolas', 9), 
                         bg='#ffffff', fg='#333333', relief='sunken', bd=2)
    status_scrollbar = tk.Scrollbar(content_frame, orient=tk.VERTICAL, command=status_text.yview)
    status_text.configure(yscrollcommand=status_scrollbar.set)
    status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    status_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    return progress_win, status_text

def update_progress(text_widget, message):
    text_widget.insert(tk.END, message + "\n")
    text_widget.see(tk.END)
    text_widget.update()

def show_completion_dialog(root, csv_file, total_freqs, direction):
    completion_win = tk.Toplevel(root)
    completion_win.title("Analysis Complete")
    completion_win.geometry("500x335")
    completion_win.configure(bg='#f0f0f0')
    completion_win.grab_set()
    
    success_frame = tk.Frame(completion_win, bg='#4CAF50', relief='raised', bd=3)
    success_frame.pack(fill=tk.X, padx=10, pady=10)
    tk.Label(success_frame, text="Analysis Complete!", font=('Arial', 16, 'bold'), 
             fg='white', bg='#4CAF50', pady=15).pack()
    
    info_frame = tk.Frame(completion_win, bg='#f0f0f0')
    info_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
    
    info_text = f"""Results saved to: {csv_file}

Frequency points processed: {total_freqs}
Analysis direction: {direction}

The CSV file contains displacement and reaction force data
for the complete frequency sweep."""
    
    tk.Label(info_frame, text=info_text, font=('Arial', 10), bg='#f0f0f0', 
             fg='#333333', justify=tk.LEFT).pack(pady=10)
    
    create_styled_button(info_frame, "Close", completion_win.destroy, style="primary", width=12).pack(pady=10)
    completion_win.wait_window()

def get_parameters(root, increment_data=None):
    """Show a professional dialog to enter analysis parameters with scrollbar."""
    global current_post_file, current_dat_file
    
    params = {}
    win = tk.Toplevel(root)
    win.title("Marc Dynamic Analysis - Parameters Configuration")
    win.geometry("685x850")
    win.configure(bg='#f0f0f0')
    win.grab_set()

    node_sets_list = get_node_sets_from_dat(current_dat_file) if current_dat_file else []

    # Use increment data from load_increments_fast
    if increment_data and 'labels' in increment_data:
        increment_labels = increment_data['labels']
        label_to_index = increment_data['label_to_index']
        has_harmonics = increment_data.get('has_harmonics', False)
    else:
        increment_labels = ['0']
        label_to_index = {'0': 0}
        has_harmonics = False

    title_frame = tk.Frame(win, bg='#d4542a', relief='raised', bd=3)
    title_frame.pack(fill=tk.X, padx=10, pady=10)
    tk.Label(title_frame, text="Marc Dynamic Analysis", font=('Arial', 16, 'bold'), 
             fg='white', bg='#d4542a', pady=8).pack()
    tk.Label(title_frame, text="Analysis Parameters Configuration", font=('Arial', 11, 'bold'), 
             fg='white', bg='#d4542a', pady=5).pack()

    # Create main container with scrollbar
    main_container = tk.Frame(win, bg='#f0f0f0')
    main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    
    # Create canvas and scrollbar
    canvas = tk.Canvas(main_container, bg='#f5f5f5', highlightthickness=0)
    scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg='#f5f5f5')
    
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    
    canvas_frame = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # Make canvas expand to fill width
    def configure_canvas(event):
        canvas.itemconfig(canvas_frame, width=event.width)
    canvas.bind('<Configure>', configure_canvas)
    
    # Mouse wheel scrolling
    def on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    canvas.bind_all("<MouseWheel>", on_mousewheel)
    
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Content frame inside scrollable area
    content_frame = tk.Frame(scrollable_frame, bg='#f5f5f5')
    content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # File Info Frame
    file_info_frame = tk.LabelFrame(content_frame, text="File Information", font=('Arial', 10, 'bold'),
                                    bg='#f5f5f5', fg='#333333', relief='groove', bd=1, padx=10, pady=5)
    file_info_frame.pack(fill=tk.X, padx=15, pady=10)
    
    tk.Label(file_info_frame, text="T16 File:", font=('Arial', 9, 'bold'), 
             bg='#f5f5f5', fg='#333333').grid(row=0, column=0, sticky='w', pady=2)
    tk.Label(file_info_frame, text=os.path.basename(current_post_file) if current_post_file else "Not loaded", 
             font=('Arial', 9), bg='#f5f5f5', fg='#006400').grid(row=0, column=1, sticky='w', padx=10)
    
    tk.Label(file_info_frame, text="Total Increments:", font=('Arial', 9, 'bold'), 
             bg='#f5f5f5', fg='#333333').grid(row=1, column=0, sticky='w', pady=2)
    tk.Label(file_info_frame, text=str(len(increment_labels)), 
             font=('Arial', 9), bg='#f5f5f5', fg='#006400').grid(row=1, column=1, sticky='w', padx=10)
    
    if has_harmonics:
        tk.Label(file_info_frame, text="(Only harmonic increments with real Marc IDs are shown)", font=('Arial', 8, 'italic'), 
                 bg='#f5f5f5', fg='#FF6600').grid(row=1, column=2, sticky='w', padx=5)

    params_frame = tk.LabelFrame(content_frame, text="Analysis Parameters", font=('Arial', 11, 'bold'),
                                bg='#f5f5f5', fg='#333333', relief='groove', bd=2, padx=15, pady=15)
    params_frame.pack(fill=tk.X, padx=15, pady=15)
    params_frame.grid_columnconfigure(1, weight=1)
    params_frame.grid_columnconfigure(3, weight=1)

    # Starting Increment
    tk.Label(params_frame, text="Starting Increment:", font=('Arial', 10, 'bold'), bg='#f5f5f5', fg='#333333')\
        .grid(row=0, column=0, columnspan=4, sticky='w', pady=(5,2))
    
    inc_combo = ttk.Combobox(params_frame, width=23, font=('Arial', 10), state='readonly')
    inc_combo['values'] = increment_labels if increment_labels else ['0']
    inc_combo.current(0)
    inc_combo.grid(row=1, column=0, columnspan=4, sticky='w', pady=(0,15))

    # Excitation Node
    tk.Label(params_frame, text="Excitation Node:", font=('Arial', 10, 'bold'), bg='#f5f5f5', fg='#333333')\
        .grid(row=2, column=0, columnspan=4, sticky='w', pady=(10,5))
    
    exc_sel_frame = tk.Frame(params_frame, bg='#f5f5f5')
    exc_sel_frame.grid(row=3, column=0, columnspan=4, sticky='w', pady=(0,5))
    
    exc_sel_var = tk.StringVar(win)
    exc_sel_var.set('Nodes')
    
    rb_exc_nodes = tk.Radiobutton(exc_sel_frame, text="Nodes", variable=exc_sel_var, value='Nodes',
                                  font=('Arial', 10, 'bold'), bg='#f5f5f5', fg='#333333',
                                  selectcolor="#FFFFFF", activebackground='#f5f5f5')
    rb_exc_nodes.pack(side=tk.LEFT, padx=(0,30))
    
    rb_exc_sets = tk.Radiobutton(exc_sel_frame, text="Node Sets", variable=exc_sel_var, value='Node Sets',
                                 font=('Arial', 10, 'bold'), bg='#f5f5f5', fg='#333333',
                                 selectcolor="#FFFFFF", activebackground='#f5f5f5')
    rb_exc_sets.pack(side=tk.LEFT)

    exc_nodes_entry = tk.Entry(params_frame, width=40, font=('Arial', 10), relief='sunken', bd=2)
    exc_nodes_entry.grid(row=4, column=0, columnspan=4, sticky='w', pady=(0,10))
    
    exc_sets_combo = ttk.Combobox(params_frame, width=37, font=('Arial', 10), state='readonly')
    exc_sets_combo['values'] = node_sets_list if node_sets_list else ['No node sets found']
    if node_sets_list:
        exc_sets_combo.current(0)
    exc_sets_combo.grid(row=4, column=0, columnspan=4, sticky='w', pady=(0,10))
    exc_sets_combo.grid_remove()

    def toggle_exc_input():
        if exc_sel_var.get() == 'Nodes':
            exc_nodes_entry.grid()
            exc_sets_combo.grid_remove()
        else:
            exc_nodes_entry.grid_remove()
            exc_sets_combo.grid()

    rb_exc_nodes.configure(command=toggle_exc_input)
    rb_exc_sets.configure(command=toggle_exc_input)

    # Reaction Force (Fixed Node)
    tk.Label(params_frame, text="Reaction Force (Fixed Node):", font=('Arial', 10, 'bold'), bg='#f5f5f5', fg='#333333')\
        .grid(row=5, column=0, columnspan=4, sticky='w', pady=(10,5))
    
    rf_sel_frame = tk.Frame(params_frame, bg='#f5f5f5')
    rf_sel_frame.grid(row=6, column=0, columnspan=4, sticky='w', pady=(0,5))
    
    rf_sel_var = tk.StringVar(win)
    rf_sel_var.set('Nodes')
    
    rb_rf_nodes = tk.Radiobutton(rf_sel_frame, text="Nodes", variable=rf_sel_var, value='Nodes',
                                 font=('Arial', 10, 'bold'), bg='#f5f5f5', fg='#333333',
                                 selectcolor="#FFFFFF", activebackground='#f5f5f5')
    rb_rf_nodes.pack(side=tk.LEFT, padx=(0,30))
    
    rb_rf_sets = tk.Radiobutton(rf_sel_frame, text="Node Sets", variable=rf_sel_var, value='Node Sets',
                                font=('Arial', 10, 'bold'), bg='#f5f5f5', fg='#333333',
                                selectcolor="#FFFFFF", activebackground='#f5f5f5')
    rb_rf_sets.pack(side=tk.LEFT)

    rf_nodes_entry = tk.Entry(params_frame, width=40, font=('Arial', 10), relief='sunken', bd=2)
    rf_nodes_entry.grid(row=7, column=0, columnspan=4, sticky='w', pady=(0,10))
    
    rf_sets_combo = ttk.Combobox(params_frame, width=37, font=('Arial', 10), state='readonly')
    rf_sets_combo['values'] = node_sets_list if node_sets_list else ['No node sets found']
    if node_sets_list:
        rf_sets_combo.current(0)
    rf_sets_combo.grid(row=7, column=0, columnspan=4, sticky='w', pady=(0,10))
    rf_sets_combo.grid_remove()

    def toggle_rf_input():
        if rf_sel_var.get() == 'Nodes':
            rf_nodes_entry.grid()
            rf_sets_combo.grid_remove()
        else:
            rf_nodes_entry.grid_remove()
            rf_sets_combo.grid()

    rb_rf_nodes.configure(command=toggle_rf_input)
    rb_rf_sets.configure(command=toggle_rf_input)

    # Direction
    tk.Label(params_frame, text="Sweep Direction:", font=('Arial', 10, 'bold'), bg='#f5f5f5', fg='#333333')\
        .grid(row=8, column=0, columnspan=4, sticky='w', pady=(10,5))
    
    direction_frame = tk.Frame(params_frame, bg='#f5f5f5')
    direction_frame.grid(row=9, column=0, columnspan=4, sticky='w', pady=(0,10))
    
    direction_var = tk.StringVar(win)
    direction_var.set('X')
    
    for direction in ['X', 'Y', 'Z']:
        tk.Radiobutton(direction_frame, text=direction, variable=direction_var, value=direction,
                       font=('Arial', 10, 'bold'), bg='#f5f5f5', fg='#333333',
                       selectcolor="#FFFFFF", activebackground='#f5f5f5').pack(side=tk.LEFT, padx=10)

    # Help section
    help_frame = tk.LabelFrame(content_frame, text="Parameter Guidelines", font=('Arial', 10, 'bold'),
                               bg='#f5f5f5', fg='#666666', relief='groove', bd=1, padx=10, pady=10)
    help_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

    help_text = """* Starting Increment: Select the harmonic increment to start the frequency sweep.
  - Only harmonic increments are listed.
  - The displayed IDs are the real Marc/Mentat IDs read from the solution file.
* Excitation Node: Choose Nodes (enter ID) or Node Sets (select from list)
* Reaction Force (Fixed Node): Choose Nodes (enter ID) or Node Sets
* Direction: X, Y, or Z
* DAT file is auto-detected (same name as T16, same folder)"""

    tk.Label(help_frame, text=help_text, font=('Arial', 9), bg='#f5f5f5', fg='#666666',
             justify=tk.LEFT, wraplength=630).pack(anchor='w')

    # Button frame (outside scrollable area, fixed at bottom)
    button_frame = tk.Frame(win, bg='#f0f0f0')
    button_frame.pack(fill=tk.X, pady=15, padx=10)

    def on_ok():
        try:
            # Get selected increment label and map to internal index
            selected_label = inc_combo.get().strip()
            if not selected_label:
                raise ValueError("Please select a valid increment")
            
            # Map label to internal index
            if selected_label in label_to_index:
                params['inc'] = label_to_index[selected_label]
            else:
                # Fallback: try to parse as integer
                params['inc'] = int(selected_label)
            
            params['direction'] = direction_var.get()

            if exc_sel_var.get() == 'Nodes':
                disp_val = exc_nodes_entry.get().strip()
                if not disp_val:
                    raise ValueError("Excitation Node (Nodes) must be filled")
                params['disp'] = disp_val
            else:
                if not node_sets_list:
                    raise ValueError("No node sets available for Excitation Node")
                set_val = exc_sets_combo.get().strip()
                if not set_val or set_val == 'No node sets found':
                    raise ValueError("Please select a valid Excitation Node Set")
                params['disp'] = set_val

            if rf_sel_var.get() == 'Nodes':
                rf_val = rf_nodes_entry.get().strip()
                if not rf_val:
                    raise ValueError("Reaction Force (Nodes) must be filled")
                params['rf_set'] = rf_val
            else:
                if not node_sets_list:
                    raise ValueError("No node sets available for Reaction Force")
                set_val = rf_sets_combo.get().strip()
                if not set_val or set_val == 'No node sets found':
                    raise ValueError("Please select a valid Reaction Force Node Set")
                params['rf_set'] = set_val

        except ValueError as e:
            messagebox.showerror("Input Error", "Please fill all fields correctly.\n\nError: " + str(e), parent=win)
            return
        
        # Unbind mousewheel before closing
        canvas.unbind_all("<MouseWheel>")
        win.destroy()

    def on_cancel():
        # Unbind mousewheel before closing
        canvas.unbind_all("<MouseWheel>")
        win.destroy()

    create_styled_button(button_frame, "Run Analysis", on_ok, style="success", width=15).pack(side=tk.LEFT, padx=10)
    create_styled_button(button_frame, "Cancel", on_cancel, style="secondary", width=12).pack(side=tk.LEFT, padx=10)

    win.protocol("WM_DELETE_WINDOW", on_cancel)
    root.wait_window(win)
    return params if params else None


def main():
    """Main application entry point"""
    try:
        import tkinter as tk
        tk_version = tk.TkVersion
    except Exception as e:
        return

    app = None
    try:
        app = MainApplication()
        app.run()
    except KeyboardInterrupt:
        if app:
            try:
                app.safe_exit()
            except:
                pass
    except Exception as e:
        pass

if __name__ == "__main__":
    main()
