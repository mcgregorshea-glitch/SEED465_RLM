import tkinter as tk
from tkinter import ttk, messagebox
import re
import os

# Optional numpy import
try:
    import numpy as np # type: ignore
    HAS_NUMPY = True
except ImportError:
    HAS_NUMPY = False

class PlotManager:
    """
    Manages the 3D (Matplotlib) and 2D (Tkinter Canvas) toolpath visualizations.
    """
    def __init__(self, gui_instance) -> None:
        self.gui = gui_instance
        self.matplotlib_imported = False
        self.fig_3d = None
        self.ax_3d = None
        self.canvas_3d = None
        self.marker_3d = None
        self._3d_was_enabled_before_live_sync = None

        # Cache for performance
        self._plot_coords_cache = None
        self._plot_cache_valid = False

        if hasattr(self.gui, 'event_bus') and self.gui.event_bus:
            self.gui.event_bus.subscribe("live_sync_changed", self._on_live_sync_changed)

    def invalidate_cache(self):
        self._plot_cache_valid = False
        self._plot_coords_cache = None

    def toggle_3d_plot(self):
        """Toggles the visibility state of the 3D plot."""
        self.gui.is_3d_plot_enabled.set(not self.gui.is_3d_plot_enabled.get())
        self.update_3d_plot_button_style()
        self.update_3d_plot_visibility()

    def update_3d_plot_button_style(self):
        """Updates the visual style of the 3D toggle button."""
        if not hasattr(self.gui, 'toggle_3d_button'): return
        if self.gui.is_3d_plot_enabled.get():
            self.gui.toggle_3d_button.config(text="DISABLE 3D VIEW", style='Custom.Toggle.On.TButton')
        else:
            self.gui.toggle_3d_button.config(text="ENABLE 3D VIEW", style='Custom.Toggle.Off.TButton')

    def update_3d_plot_visibility(self):
        """Refreshes the 3D plot tab content based on enable state."""
        if not hasattr(self.gui, 'plot_container_frame'): return

        for widget in self.gui.plot_container_frame.winfo_children():
            widget.destroy()

        if self.gui.is_3d_plot_enabled.get():
            self.create_3d_plot_widgets(self.gui.plot_container_frame)
        else:
            self.ax_3d = None; self.canvas_3d = None; self.fig_3d = None; self.marker_3d = None
            disabled_label = ttk.Label(
                self.gui.plot_container_frame,
                text="3D Plot disabled for performance.",
                justify=tk.CENTER,
                font=self.gui.FONT_MONO,
                foreground=self.gui.COLOR_TEXT_SECONDARY
            )
            disabled_label.pack(fill=tk.BOTH, expand=True, pady=20)
            self.invalidate_cache()

    def _on_live_sync_changed(self, data):
        """Disables the 3D toolpath plot while live sync is active to reduce GPU load."""
        active = data.get("active", False)
        if active:
            self._3d_was_enabled_before_live_sync = self.gui.is_3d_plot_enabled.get()
            if self.gui.is_3d_plot_enabled.get():
                self.gui.is_3d_plot_enabled.set(False)
                self.update_3d_plot_button_style()
                self.update_3d_plot_visibility()
        else:
            if self._3d_was_enabled_before_live_sync is not None:
                self.gui.is_3d_plot_enabled.set(self._3d_was_enabled_before_live_sync)
                self._3d_was_enabled_before_live_sync = None
                self.update_3d_plot_button_style()
                self.update_3d_plot_visibility()

    def create_3d_plot_widgets(self, parent):
        """Instantiates matplotlib widgets for 3D visualization."""
        plot_frame = ttk.Frame(parent, style="Panel.TFrame")
        plot_frame.pack(fill=tk.BOTH, expand=True)

        def deferred_load():
            try:
                from matplotlib.figure import Figure
                from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
                
                self.matplotlib_imported = True
                
                for widget in plot_frame.winfo_children():
                    widget.destroy()

                self.fig_3d = Figure(figsize=(5, 4), dpi=100, facecolor=self.gui.COLOR_PANEL_BG)
                self.ax_3d = self.fig_3d.add_subplot(111, projection='3d')
                
                self.canvas_3d = FigureCanvasTkAgg(self.fig_3d, plot_frame)
                self.gui.canvas_3d = self.canvas_3d # Keep reference in GUI too
                self.canvas_3d.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
                
                self.create_view_controls(self.gui._3d_control_bar)
                self.draw_3d_toolpath()

            except ImportError:
                error_label = ttk.Label(plot_frame, text="3D view requires 'matplotlib'.\n\nPlease install it by running:\npip install matplotlib", foreground=self.gui.COLOR_ACCENT_AMBER, justify=tk.CENTER, font=self.gui.FONT_MONO)
                error_label.pack(fill=tk.BOTH, expand=True)

        if not self.matplotlib_imported:
            loading_label = ttk.Label(plot_frame, text="Loading 3D visualization...", justify=tk.CENTER, font=self.gui.FONT_MONO, foreground=self.gui.COLOR_TEXT_SECONDARY)
            loading_label.pack(fill=tk.BOTH, expand=True)
            self.gui.root.after(50, deferred_load)
        else:
            deferred_load()

    def create_view_controls(self, parent):
        """Creates toolbar for 3D view controls."""
        if not self.matplotlib_imported: return

        controls_frame = ttk.Frame(parent, style="Panel.TFrame")
        controls_frame.pack(side=tk.RIGHT, padx=(10, 0))

        ttk.Button(controls_frame, text="↑", command=lambda: self.rotate_view(elev_change=15), style='ViewCube.TButton', width=2).pack(side=tk.LEFT, padx=1)
        ttk.Button(controls_frame, text="↓", command=lambda: self.rotate_view(elev_change=-15), style='ViewCube.TButton', width=2).pack(side=tk.LEFT, padx=1)
        ttk.Button(controls_frame, text="←", command=lambda: self.rotate_view(azim_change=15), style='ViewCube.TButton', width=2).pack(side=tk.LEFT, padx=1)
        ttk.Button(controls_frame, text="→", command=lambda: self.rotate_view(azim_change=-15), style='ViewCube.TButton', width=2).pack(side=tk.LEFT, padx=1)

        ttk.Separator(controls_frame, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=4, pady=2)

        ttk.Button(controls_frame, text="Top", command=lambda: self.set_view(elev=90, azim=-90), style="Segment.TButton", width=4).pack(side=tk.LEFT, padx=1)
        ttk.Button(controls_frame, text="Front", command=lambda: self.set_view(elev=0, azim=-90), style="Segment.TButton", width=5).pack(side=tk.LEFT, padx=1)
        ttk.Button(controls_frame, text="Iso", command=lambda: self.set_view(elev=30, azim=-60), style="Segment.TButton", width=3).pack(side=tk.LEFT, padx=1)

    def set_view(self, elev, azim):
        if not self.matplotlib_imported or not self.ax_3d: return
        self.ax_3d.view_init(elev=elev, azim=azim)
        self.canvas_3d.draw()

    def rotate_view(self, elev_change=0, azim_change=0):
        if not self.matplotlib_imported or not self.ax_3d: return
        new_elev = max(-90, min(90, self.ax_3d.elev + elev_change))
        new_azim = self.ax_3d.azim + azim_change
        self.set_view(new_elev, new_azim)

    def _are_collinear(self, p1, p2, p3, tolerance=1e-6):
        x1, y1, z1 = p1; x2, y2, z2 = p2; x3, y3, z3 = p3
        area = x1 * (y2 - y3) + x2 * (y3 - y1) + x3 * (y1 - y2)
        if abs(area) > tolerance: return False
        area_xz = x1 * (z2 - z3) + x2 * (z3 - z1) + x3 * (z1 - z2)
        if abs(area_xz) > tolerance: return False
        area_yz = y1 * (z2 - z3) + y2 * (z3 - z1) + y3 * (z1 - z2)
        if abs(area_yz) > tolerance: return False
        return True

    def build_plot_coordinates(self, toolpath_by_layer, ordered_z_values, move_to_layer_map):
        if self._plot_cache_valid and self._plot_coords_cache:
            return self._plot_coords_cache
        
        all_points = []
        if ordered_z_values:
            try:
                first_z = next(iter(sorted(toolpath_by_layer.keys())))
                first_segment = toolpath_by_layer[first_z][0]
                all_points.append((first_segment[0][0], first_segment[0][1], first_z))
                for z, idx in move_to_layer_map:
                    segment = toolpath_by_layer[z][idx]
                    all_points.append((segment[1][0], segment[1][1], z))
            except (IndexError, StopIteration):
                all_points = []

        if len(all_points) < 3:
            simplified_points = all_points
        else:
            simplified_points = [all_points[0]]
            for i in range(1, len(all_points) - 1):
                if not self._are_collinear(simplified_points[-1], all_points[i], all_points[i+1]):
                    simplified_points.append(all_points[i])
            simplified_points.append(all_points[-1])

        if not simplified_points:
            self._plot_coords_cache = ([], [], [])
            self._plot_cache_valid = True
            return self._plot_coords_cache

        if HAS_NUMPY:
            coords_array = np.array(simplified_points, dtype=np.float32)
            x, y, z = coords_array[:, 0], coords_array[:, 1], coords_array[:, 2]
        else:
            x = [p[0] for p in simplified_points]; y = [p[1] for p in simplified_points]; z = [p[2] for p in simplified_points]
        
        self._plot_coords_cache = (x, y, z)
        self._plot_cache_valid = True
        return self._plot_coords_cache

    def draw_2d_toolpath(self, canvas):
        """Draws the 2D toolpath on the provided Tkinter canvas."""
        if not canvas or not self.gui.is_2d_plot_enabled.get():
            return
        
        canvas.delete("toolpath")
        x, y, z = self.build_plot_coordinates(self.gui.toolpath_by_layer, self.gui.ordered_z_values, self.gui.move_to_layer_map)
        
        if len(x) < 2:
            return
            
        w, h = canvas.winfo_width(), canvas.winfo_height()
        if w < 10: w, h = 215, 215
        x_max, y_max = self.gui.PRINTER_BOUNDS['x_max'], self.gui.PRINTER_BOUNDS['y_max']
        
        points = []
        for i in range(len(x)):
            px = (x[i] / x_max) * w
            py = (1 - (y[i] / y_max)) * h
            points.extend([px, py])
            
        if len(points) >= 4:
            canvas.create_line(points, fill=self.gui.COLOR_ACCENT_CYAN, width=1, tags="toolpath")

    def draw_3d_toolpath(self):
        if not self.ax_3d or not self.gui.is_3d_plot_enabled.get(): return
        curr_elev, curr_azim = self.ax_3d.elev, self.ax_3d.azim
        self.ax_3d.clear()
        self.marker_3d = None
        self.style_3d_plot(elev=curr_elev, azim=curr_azim)
        x, y, z = self.build_plot_coordinates(self.gui.toolpath_by_layer, self.gui.ordered_z_values, self.gui.move_to_layer_map)
        if len(x) > 1:
            self.ax_3d.plot(x, y, z, color=self.gui.COLOR_ACCENT_CYAN, linewidth=1.2, alpha=self.gui.toolpath_3d_opacity_var.get())
        bounds = self.gui.PRINTER_BOUNDS
        self.ax_3d.set_xlim(bounds['x_min'], bounds['x_max']); self.ax_3d.set_ylim(bounds['y_min'], bounds['y_max']); self.ax_3d.set_zlim(bounds['z_min'], bounds['z_max'])
        self.update_3d_position_marker()
        self.canvas_3d.draw()

    def style_3d_plot(self, elev=None, azim=None):
        ax = self.ax_3d
        ax.grid(False)
        ax.set_facecolor(self.gui.COLOR_PANEL_BG)
        ax.xaxis.set_pane_color((0, 0, 0, 0)); ax.yaxis.set_pane_color((0, 0, 0, 0)); ax.zaxis.set_pane_color((0, 0, 0, 0))
        ax.set_xlabel('X (mm)', color=self.gui.COLOR_TEXT_SECONDARY, fontsize=8)
        ax.set_ylabel('Y (mm)', color=self.gui.COLOR_TEXT_SECONDARY, fontsize=8)
        ax.set_zlabel('Z (mm)', color=self.gui.COLOR_TEXT_SECONDARY, fontsize=8)
        for axis in [ax.xaxis, ax.yaxis, ax.zaxis]:
            axis.set_tick_params(colors=self.gui.COLOR_TEXT_SECONDARY, labelsize=7)
        if elev is not None and azim is not None: ax.view_init(elev=elev, azim=azim)
        else: ax.view_init(elev=20, azim=-45)

    def update_3d_position_marker(self):
        if not self.ax_3d: return
        x = self.gui.last_cmd_abs_x or self.gui.PRINTER_BOUNDS['x_min']
        y = self.gui.last_cmd_abs_y or self.gui.PRINTER_BOUNDS['y_min']
        z = self.gui.last_cmd_abs_z or self.gui.PRINTER_BOUNDS['z_min']
        if self.marker_3d:
            self.marker_3d.remove()
        self.marker_3d = self.ax_3d.scatter([x], [y], [z], color=self.gui.COLOR_ACCENT_RED, s=40, depthshade=False)
        if self.canvas_3d:
            self.canvas_3d.draw_idle()
