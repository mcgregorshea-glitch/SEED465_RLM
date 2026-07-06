import tkinter as tk
from tkinter import ttk

class EventBus:
    """Simple pub/sub event bus."""
    def __init__(self):
        self._subscribers = {}

    def subscribe(self, event_type, handler):
        if event_type not in self._subscribers:
            self._subscribers[event_type] = []
        self._subscribers[event_type].append(handler)

    def publish(self, event_type, data):
        if event_type in self._subscribers:
            for handler in self._subscribers[event_type]:
                handler(data)
# If you change the physical dimensions of the test fixture, update these values.
# PRINTER_LIMITS is used by the Generator to prevent users from designing 
# scan paths that exceed the Ender 3's physical travel.
PRINTER_LIMITS = {
    'x': 110.0,      # Center-relative limit (220mm total / 2)
    'y': 110.0,      # Center-relative limit (220mm total / 2)
    'z_max': 128.0,
    'z_min': 0.0,
    'rot_max': 90.0,
    'rot_min': -90.0,
}

# PRINTER_BOUNDS is used by the Sender for runtime validation.
# It uses absolute machine coordinates, matching the printer firmware's travel.
PRINTER_BOUNDS = {
    'x_min': 0,   'x_max': 220,
    'y_min': 0,   'y_max': 220,
    'z_min': 0,   'z_max': 128,
    'rot_min': -90, 'rot_max': 90,  # Rotation axis ±90° physical limit
}

# --- Color Palette ---
COLOR_BG = "#0a0e14"
COLOR_PANEL_BG = "#161b22"
COLOR_BORDER = "#30363d"
COLOR_TEXT_PRIMARY = "#e6edf3"
COLOR_TEXT_SECONDARY = "#7d8590"
COLOR_ACCENT_CYAN = "#00d4ff"
COLOR_ACCENT_PURPLE = "#a371f7"
COLOR_ACCENT_GREEN = "#3fb950"
COLOR_ACCENT_AMBER = "#ffa657"
COLOR_ACCENT_RED = "#ff4444"
COLOR_BLACK = "#000000"
COLOR_GREY_COMPLETED = "#484f58"

# --- Fonts ---
FONT_HEADER = ("Orbitron", 13)
FONT_BODY = ("Inter", 11)
FONT_BODY_SMALL = ("Inter", 9)
FONT_BODY_BOLD = ("Inter", 11, "bold")
FONT_BODY_BOLD_LARGE = ("Inter", 20, "bold")
FONT_MONO = ("JetBrains Mono", 10)
FONT_MONO_LARGE = ('JetBrains Mono', 11, 'bold')
FONT_DRO = ("Space Mono", 16, "bold")
FONT_TERMINAL = ("JetBrains Mono", 10)

def setup_global_styling(root):
    """Configures the global ttk styles based on the sleak dark theme."""
    style = ttk.Style()
    
    # Use 'clam' as a base as it is highly customizable
    if 'clam' in style.theme_names():
        style.theme_use('clam')
        
    style.configure('.',
                    background=COLOR_PANEL_BG,
                    foreground=COLOR_TEXT_PRIMARY,
                    fieldbackground=COLOR_BLACK,
                    bordercolor=COLOR_BORDER,
                    lightcolor=COLOR_BORDER,
                    darkcolor=COLOR_BORDER,
                    font=FONT_BODY)
                    
    style.map('.',
              background=[('disabled', COLOR_PANEL_BG), ('active', COLOR_PANEL_BG)],
              foreground=[('disabled', COLOR_TEXT_SECONDARY)],
              bordercolor=[('focus', COLOR_ACCENT_CYAN), ('active', COLOR_BORDER)],
              fieldbackground=[('disabled', COLOR_PANEL_BG)])

    # Frames
    style.configure('TFrame', background=COLOR_BG)
    style.configure('Dark.TFrame', background=COLOR_BG)
    style.configure('Panel.TFrame', background=COLOR_PANEL_BG)
    style.configure('Header.TFrame', background=COLOR_PANEL_BG, bordercolor=COLOR_BORDER, borderwidth=1, relief='solid')
    style.configure('Footer.TFrame', background=COLOR_BLACK, bordercolor=COLOR_BORDER, borderwidth=1, relief='solid')
    style.configure('Black.TFrame', background=COLOR_BLACK)
    
    # LabelFrames
    style.configure('TLabelframe',
                    background=COLOR_PANEL_BG,
                    bordercolor=COLOR_BORDER,
                    borderwidth=1,
                    relief=tk.SOLID,
                    padding=16)
    style.configure('TLabelframe.Label',
                    background=COLOR_PANEL_BG,
                    foreground=COLOR_TEXT_SECONDARY,
                    font=("Rajdhani", 13, "bold"),
                    padding=(10, 5))
                    
    style.configure('Card.TLabelframe',
                    background=COLOR_PANEL_BG,
                    bordercolor=COLOR_BORDER,
                    borderwidth=1,
                    relief=tk.SOLID,
                    padding=12)
    style.configure('Card.TLabelframe.Label',
                    background=COLOR_PANEL_BG,
                    foreground=COLOR_ACCENT_CYAN,
                    font=('Inter', 10, 'bold'))

    # Color-coded LabelFrames
    for name, color in [('Yellow', COLOR_ACCENT_AMBER), ('Green', COLOR_ACCENT_GREEN), ('Purple', "#c4c1ff"), ('Cyan', COLOR_ACCENT_CYAN)]:
        style.configure(f'{name}.TLabelframe',
                        background=COLOR_PANEL_BG,
                        bordercolor=color,
                        borderwidth=2,
                        relief=tk.SOLID)
        style.configure(f'{name}.TLabelframe.Label',
                        background=COLOR_PANEL_BG,
                        foreground=color,
                        font=("Rajdhani", 13, "bold"))

    # Labels
    style.configure('TLabel', background=COLOR_PANEL_BG, foreground=COLOR_TEXT_PRIMARY, font=FONT_BODY)
    style.configure('Secondary.TLabel', background=COLOR_PANEL_BG, foreground=COLOR_TEXT_SECONDARY, font=FONT_BODY_SMALL)
    style.configure('Filename.TLabel', background=COLOR_PANEL_BG, foreground=COLOR_ACCENT_CYAN, font=FONT_MONO)
    style.configure('Header.TLabel', background=COLOR_PANEL_BG, font=FONT_BODY)
    style.configure('Footer.TLabel', background=COLOR_BLACK, foreground=COLOR_TEXT_SECONDARY, font=FONT_MONO)
    style.configure('Filepath.TLabel', background=COLOR_PANEL_BG, foreground=COLOR_TEXT_SECONDARY, font=FONT_BODY_SMALL)
    
    style.configure('DRO.TLabel', font=FONT_MONO, padding=(5, 5), background=COLOR_BLACK, foreground=COLOR_TEXT_SECONDARY, borderwidth=1, relief='sunken', anchor='w')
    style.configure('Red.DRO.TLabel', font=FONT_DRO, width=8, padding=(5, 5), background=COLOR_BLACK, foreground=COLOR_ACCENT_RED, borderwidth=0, relief='flat', anchor='e')
    style.configure('Blue.DRO.TLabel', font=FONT_DRO, width=8, padding=(5, 5), background=COLOR_BLACK, foreground=COLOR_ACCENT_AMBER, borderwidth=0, relief='flat', anchor='e')

    # Buttons
    style.configure('TButton', background=COLOR_PANEL_BG, foreground=COLOR_TEXT_PRIMARY, bordercolor=COLOR_BORDER, borderwidth=1, relief=tk.SOLID, padding=(12, 8), font=FONT_BODY)
    style.map('TButton', background=[('active', '#2c333e'), ('pressed', COLOR_BLACK)], foreground=[('active', COLOR_ACCENT_CYAN)], bordercolor=[('active', COLOR_ACCENT_CYAN)])

    style.configure('Primary.TButton', background=COLOR_ACCENT_CYAN, foreground=COLOR_BLACK, padding=(12, 10), font=FONT_BODY_BOLD)
    style.map('Primary.TButton', background=[('active', '#00eaff'), ('pressed', COLOR_ACCENT_CYAN)], foreground=[('active', COLOR_BLACK), ('pressed', COLOR_BLACK)], bordercolor=[('active', COLOR_ACCENT_CYAN)])

    style.configure('Danger.TButton', background=COLOR_ACCENT_RED, foreground=COLOR_TEXT_PRIMARY, font=FONT_BODY_BOLD)
    style.map('Danger.TButton', background=[('active', '#ff6666'), ('pressed', COLOR_ACCENT_RED)], bordercolor=[('active', COLOR_ACCENT_RED)])

    style.configure('Amber.TButton', background=COLOR_ACCENT_AMBER, foreground=COLOR_BLACK, font=FONT_BODY_BOLD)
    style.map('Amber.TButton', background=[('active', '#ffc080'), ('pressed', COLOR_ACCENT_AMBER)], bordercolor=[('active', COLOR_ACCENT_AMBER)])

    style.configure('Purple.TButton', background=COLOR_ACCENT_PURPLE, foreground=COLOR_BLACK, font=FONT_BODY_BOLD)
    style.map('Purple.TButton', background=[('active', '#b890ff'), ('pressed', COLOR_ACCENT_PURPLE)], bordercolor=[('active', COLOR_ACCENT_PURPLE)])

    # Color-coded Ring Buttons
    for name, color in [('Yellow', COLOR_ACCENT_AMBER), ('Green', COLOR_ACCENT_GREEN), ('Purple', COLOR_ACCENT_PURPLE), ('Cyan', COLOR_ACCENT_CYAN)]:
        style.configure(f'{name}Ring.TButton', bordercolor=color, borderwidth=2)
        style.map(f'{name}Ring.TButton', 
                  foreground=[('active', color)],
                  bordercolor=[('active', color), ('pressed', COLOR_BLACK)])

    style.configure('Small.TButton', padding=(5, 2), font=FONT_BODY_SMALL)

    style.configure('Segment.TButton', background=COLOR_PANEL_BG, foreground=COLOR_TEXT_SECONDARY, padding=(10, 5), font=FONT_BODY_SMALL)
    style.map('Segment.TButton', background=[('active', '#2c333e'), ('pressed', COLOR_BLACK)], foreground=[('active', COLOR_ACCENT_CYAN)])
    style.configure('Segment.Active.TButton', background=COLOR_ACCENT_CYAN, foreground=COLOR_BLACK, padding=(10, 5), font=FONT_BODY_SMALL)
    style.map('Segment.Active.TButton', background=[('active', COLOR_ACCENT_CYAN), ('pressed', COLOR_ACCENT_CYAN)], foreground=[('active', COLOR_BLACK), ('pressed', COLOR_BLACK)])

    style.configure('Jog.TButton', font=FONT_BODY_BOLD, width=5, padding=(10, 10))
    style.configure('JogIcon.TButton', font=("Inter", 18, "bold"), width=5, padding=(4, 4))
    style.configure('Home.TButton', font=FONT_BODY_BOLD_LARGE, width=5, padding=(4, 4))
    
    style.configure('ViewCube.TButton', padding=(2, 2), font=("Inter", 12), width=2)
    style.map('ViewCube.TButton', background=[('active', '#2c333e'), ('pressed', COLOR_BLACK)], foreground=[('active', COLOR_ACCENT_CYAN)])

    # Toggle Button Styles
    custom_font = ("Rajdhani", 9, "bold")
    padding_toggle = (5, 3)
    style.configure('Custom.Toggle.Off.TButton', background=COLOR_PANEL_BG, foreground=COLOR_TEXT_SECONDARY, bordercolor=COLOR_BORDER, font=custom_font, padding=padding_toggle)
    style.map('Custom.Toggle.Off.TButton', bordercolor=[('active', COLOR_ACCENT_CYAN)], foreground=[('active', COLOR_ACCENT_CYAN)])
    style.configure('Custom.Toggle.On.TButton', background=COLOR_PANEL_BG, foreground=COLOR_ACCENT_CYAN, bordercolor=COLOR_ACCENT_CYAN, font=custom_font, padding=padding_toggle)
    style.map('Custom.Toggle.On.TButton', bordercolor=[('active', COLOR_ACCENT_CYAN)])

    # Inputs (Entry)
    style.configure('TEntry',
                    fieldbackground=COLOR_BLACK,
                    foreground=COLOR_ACCENT_CYAN,
                    bordercolor=COLOR_BORDER,
                    insertcolor=COLOR_ACCENT_CYAN,
                    borderwidth=1,
                    relief=tk.SOLID,
                    padding=6,
                    font=FONT_MONO)
    style.map('TEntry',
              fieldbackground=[('focus', COLOR_BLACK)],
              foreground=[('focus', COLOR_ACCENT_CYAN)],
              bordercolor=[('focus', COLOR_ACCENT_CYAN)])
              
    # Combobox
    style.configure('TCombobox',
                    fieldbackground=COLOR_BLACK,
                    foreground=COLOR_ACCENT_CYAN,
                    bordercolor=COLOR_BORDER,
                    arrowcolor=COLOR_ACCENT_CYAN,
                    background=COLOR_BLACK,
                    padding=6,
                    font=FONT_MONO)
    
    style.map('TCombobox', 
              fieldbackground=[('readonly', COLOR_BLACK), ('focus', COLOR_BLACK), ('disabled', COLOR_PANEL_BG)],
              foreground=[('readonly', COLOR_ACCENT_CYAN), ('focus', COLOR_ACCENT_CYAN), ('disabled', COLOR_TEXT_SECONDARY)],
              bordercolor=[('focus', COLOR_ACCENT_CYAN)])
    
    # Global overrides for dropdown listboxes (The 'Legacy' Tk quirk)
    # We use both lowercase and capitalized keys with wildcard separators to ensure coverage across all OS variants
    listbox_overrides = [
        ('background', COLOR_BLACK),
        ('foreground', COLOR_ACCENT_CYAN),
        ('selectBackground', COLOR_ACCENT_CYAN),
        ('selectForeground', COLOR_BLACK),
        ('Background', COLOR_BLACK),
        ('Foreground', COLOR_ACCENT_CYAN),
        ('SelectBackground', COLOR_ACCENT_CYAN),
        ('SelectForeground', COLOR_BLACK)
    ]
    
    for attr, val in listbox_overrides:
        root.option_add(f'*TCombobox*Listbox*{attr}', val)
        root.option_add(f'*Combobox*Listbox*{attr}', val)
        root.option_add(f'*Listbox*{attr}', val)

    root.option_add('*TCombobox*Listbox.font', FONT_MONO)
    root.option_add('*TCombobox*Listbox.borderWidth', 0)
    root.option_add('*Listbox.font', FONT_MONO)
    root.option_add('*Listbox.borderWidth', 0)

    # Checkbutton & Radiobutton
    style.configure('TCheckbutton', background=COLOR_PANEL_BG, foreground=COLOR_TEXT_SECONDARY, font=FONT_BODY_SMALL)
    style.map('TCheckbutton', background=[('active', COLOR_PANEL_BG)], foreground=[('active', COLOR_ACCENT_CYAN), ('selected', COLOR_ACCENT_CYAN)])
    style.configure('TRadiobutton', background=COLOR_PANEL_BG, foreground=COLOR_TEXT_SECONDARY, font=FONT_BODY_SMALL)
    style.map('TRadiobutton', background=[('active', COLOR_PANEL_BG)], foreground=[('active', COLOR_ACCENT_CYAN), ('selected', COLOR_ACCENT_CYAN)])

    # Scrollbars, Progress Bars, etc.
    style.configure('Vertical.TScrollbar', background=COLOR_BORDER, troughcolor=COLOR_BG, bordercolor=COLOR_BG, arrowcolor=COLOR_TEXT_PRIMARY, relief=tk.FLAT, arrowsize=14)
    style.map('Vertical.TScrollbar', background=[('active', COLOR_ACCENT_CYAN), ('!active', COLOR_BORDER)], troughcolor=[('active', COLOR_BG), ('!active', COLOR_BG)])

    style.configure('TScrollbar', background=COLOR_BORDER, troughcolor=COLOR_BG, bordercolor=COLOR_BG, arrowcolor=COLOR_TEXT_PRIMARY, relief=tk.FLAT, arrowsize=14)
    style.map('TScrollbar', background=[('active', COLOR_ACCENT_CYAN), ('!active', COLOR_BORDER)], troughcolor=[('active', COLOR_BG), ('!active', COLOR_BG)])

    style.configure('TProgressbar', troughcolor=COLOR_BLACK, background=COLOR_ACCENT_CYAN, bordercolor=COLOR_BORDER, borderwidth=1, relief=tk.SOLID)
    style.configure('Sash', background=COLOR_BG, sashthickness=6, relief=tk.FLAT)
    style.map('Sash', background=[('active', COLOR_ACCENT_CYAN)])
    
    # Notebook
    style.configure('TNotebook', tabposition='n', borderwidth=0, background=COLOR_BG)
    style.configure('TNotebook.Tab', background=COLOR_PANEL_BG, foreground=COLOR_TEXT_SECONDARY, padding=[12, 6], font=FONT_BODY)
    style.map('TNotebook.Tab', 
              background=[('selected', COLOR_BG), ('active', COLOR_BORDER)], 
              foreground=[('selected', COLOR_ACCENT_CYAN), ('active', COLOR_TEXT_PRIMARY)])
