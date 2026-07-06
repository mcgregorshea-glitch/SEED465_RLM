import tkinter as tk
from tkinter import ttk
from utils import PRINTER_LIMITS

class PatternPreview(ttk.Frame):
    """
    Component for visualizing the scan volume and statistics.
    Encapsulates the 2D wireframe canvas and textual metrics readout.
    """
    def __init__(self, parent, controller):
        super().__init__(parent, style='Dark.TFrame')
        self.c = controller
        self._setup_ui()

    def _setup_ui(self):
        title = ttk.Label(self, text="SCAN VOLUME PREVIEW",
                            font=('Rajdhani', 16, 'bold'),
                            foreground=self.c.COLOR_ACCENT_CYAN, background=self.c.COLOR_BG)
        title.pack(side=tk.TOP, pady=(0, 10), anchor=tk.W)

        # Bottom Section: Statistics Readout
        stats_frame = ttk.LabelFrame(self, text="Scan Statistics", style='Card.TLabelframe', padding=0)
        stats_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))

        self.c.stats_text = tk.Text(stats_frame, height=9, width=50,
                                    state='disabled', wrap=tk.WORD,
                                    bg=self.c.COLOR_BLACK,
                                    fg=self.c.COLOR_ACCENT_GREEN,
                                    font=self.c.FONT_MONO,
                                    relief=tk.FLAT,
                                    bd=0,
                                    padx=12,
                                    pady=10
                                    )
        self.c.stats_text.pack(fill=tk.BOTH, expand=True)

        # Style tags for statistics formatting
        self.c.stats_text.tag_configure('header', foreground=self.c.COLOR_ACCENT_CYAN, font=self.c.FONT_MONO_LARGE)
        self.c.stats_text.tag_configure('value', foreground=self.c.COLOR_ACCENT_GREEN, font=self.c.FONT_MONO_LARGE)
        self.c.stats_text.tag_configure('warning', foreground=self.c.COLOR_ACCENT_RED, font=(self.c.FONT_MONO[0], self.c.FONT_MONO[1], 'bold'))
        self.c.stats_text.tag_configure('amber_warning', foreground=self.c.COLOR_ACCENT_AMBER, font=(self.c.FONT_MONO[0], self.c.FONT_MONO[1], 'bold'))
        self.c.stats_text.tag_configure('success', foreground=self.c.COLOR_ACCENT_GREEN, font=(self.c.FONT_MONO[0], self.c.FONT_MONO[1], 'bold'))
        self.c.stats_text.tag_configure('label', foreground=self.c.COLOR_TEXT_SECONDARY, font=self.c.FONT_MONO)
        self.c.stats_text.tag_configure('separator', foreground=self.c.COLOR_BORDER, font=self.c.FONT_MONO)

        # Top Section: 3D Visualization Canvas
        canvas_frame = ttk.Frame(self, style='Dark.TFrame')
        canvas_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 0))

        self.c.preview_canvas = tk.Canvas(canvas_frame, bg=self.c.COLOR_BLACK, highlightthickness=1,
                                            highlightbackground=self.c.COLOR_BORDER)
        self.c.preview_canvas.pack(fill=tk.BOTH, expand=True)
        self.c.preview_canvas.bind("<Configure>", self.c._on_canvas_resize)

        self.draw_preview_diagram(None, [], 0)

    def draw_preview_diagram(self, params, bounds_warnings, warning_level=0):
        """Renders the 2D oblique wireframe."""
        canvas = self.c.preview_canvas
        canvas.delete("all")

        canvas_w = canvas.winfo_width()
        canvas_h = canvas.winfo_height()

        if canvas_w <= 1 or canvas_h <= 1:
            canvas.after(50, self.draw_preview_diagram, params, bounds_warnings, warning_level)
            return

        if params is None:
            canvas.create_text(canvas_w / 2, canvas_h / 2, text="Waiting for parameters...", fill=self.c.COLOR_TEXT_SECONDARY, font=self.c.FONT_BODY)
            params = {'x_min': 0, 'x_max': 0, 'y_min': 0, 'y_max': 0, 'z_min': 0, 'z_max': 0}
        
        # Oblique Projection Logic
        pl = PRINTER_LIMITS
        total_min_x, total_max_x = min(params['x_min'], -pl['x']), max(params['x_max'], pl['x'])
        total_min_y, total_max_y = min(params['y_min'], -pl['y']), max(params['y_max'], pl['y'])
        total_min_z, total_max_z = min(params['z_min'], pl['z_min']), max(params['z_max'], pl['z_max'])

        total_x_rng = max(1, total_max_x - total_min_x)
        total_y_rng = max(1, total_max_y - total_min_y)
        total_z_rng = max(1, total_max_z - total_min_z)

        pad, oblique = 40, 0.4
        total_w_units = total_x_rng + (total_y_rng * oblique)
        total_h_units = total_z_rng + (total_y_rng * oblique)
        scale = min((canvas_w - 2 * pad) / total_w_units, (canvas_h - 2 * pad) / total_h_units)

        def project(x, y, z):
            x_pct, y_pct, z_pct = (x - total_min_x)/total_x_rng, (y - total_min_y)/total_y_rng, (z - total_min_z)/total_z_rng
            sw, sh, sd = total_x_rng * scale, total_z_rng * scale, total_y_rng * scale * oblique
            x_start, y_start = (canvas_w - (sw + sd)) / 2, (canvas_h - (sh + sd)) / 2
            return (x_start + x_pct * sw + y_pct * sd, y_start + (1 - z_pct) * sh + y_pct * sd)

        # Draw Safety Limits
        pts = [project(x, y, z) for z in [pl['z_min'], pl['z_max']] for y in [-pl['y'], pl['y']] for x in [-pl['x'], pl['x']]]
        edges = [(0,1), (1,3), (3,2), (2,0), (4,5), (5,7), (7,6), (6,4), (0,4), (1,5), (2,6), (3,7)]
        for i1, i2 in edges: canvas.create_line(pts[i1], pts[i2], fill=self.c.COLOR_ACCENT_RED, dash=(4, 4), width=1)

        # Draw Pattern Volume
        if params['x_max'] != params['x_min']:
            ppts = [project(x, y, z) for z in [params['z_min'], params['z_max']] for y in [params['y_min'], params['y_max']] for x in [params['x_min'], params['x_max']]]
            for i1, i2 in edges: canvas.create_line(ppts[i1], ppts[i2], fill=self.c.COLOR_ACCENT_CYAN, width=2)

        # Labels & Origin
        if total_min_x <= 0 <= total_max_x and total_min_y <= 0 <= total_max_y:
            ox, oy = project(0,0,0)
            canvas.create_oval(ox-5, oy-5, ox+5, oy+5, fill=self.c.COLOR_ACCENT_RED)
