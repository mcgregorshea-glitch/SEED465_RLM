import tkinter as tk
from tkinter import ttk
import math

class StatePanel(tk.Frame):
    """
    Displays Z-axis height and a rotating Extruder spinner.
    """
    def __init__(self, master):
        super().__init__(master)
        
        # Z-Axis UI
        self.z_frame = tk.LabelFrame(self, text="Z-Height (mm)", padx=10, pady=10)
        self.z_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        self.z_bar = ttk.Progressbar(self.z_frame, orient=tk.VERTICAL, length=200, maximum=250)
        self.z_bar.pack(side=tk.LEFT)
        
        self.z_label = tk.Label(self.z_frame, text="0.00", font=("Consolas", 12))
        self.z_label.pack(side=tk.LEFT, padx=5)
        
        # Extruder UI
        self.e_frame = tk.LabelFrame(self, text="Extruder Spinner", padx=10, pady=10)
        self.e_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        self.e_canvas = tk.Canvas(self.e_frame, width=100, height=100, bg="#2d2d2d", highlightthickness=0)
        self.e_canvas.pack()
        
        # Draw a simple gear/spinner shape
        self.gear_id = self._draw_gear(50, 50, 40)
        
        self.e_dist_label = tk.Label(self.e_frame, text="E: 0.00 mm", font=("Consolas", 10))
        self.e_dist_label.pack(pady=5)
        
        self.current_e = 0.0

    def _draw_gear(self, cx, cy, r):
        """Draws a gear-like shape on the canvas."""
        # Simple cross/spoke design
        gear_parts = []
        gear_parts.append(self.e_canvas.create_oval(cx-r, cy-r, cx+r, cy+r, outline="white", width=2))
        gear_parts.append(self.e_canvas.create_line(cx-r, cy, cx+r, cy, fill="white", width=2))
        gear_parts.append(self.e_canvas.create_line(cx, cy-r, cx, cy+r, fill="white", width=2))
        return gear_parts

    def update_state(self, z: float, e: float):
        """Updates the Z-bar and rotates the E-spinner."""
        # Update Z
        self.z_bar['value'] = z
        self.z_label.config(text=f"{z:.2f}")
        
        # Update E
        de = e - self.current_e
        self.current_e = e
        self.e_dist_label.config(text=f"E: {e:.2f} mm")
        
        # Rotate "gear" based on E (1mm = 36 degrees for visualization)
        angle = (e % 10) * 36  # Simplistic rotation
        self._rotate_gear(50, 50, 40, angle)

    def _rotate_gear(self, cx, cy, r, angle_deg):
        """Rotates the gear spokes."""
        rad = math.radians(angle_deg)
        
        # Spoke 1
        x1 = cx + r * math.cos(rad)
        y1 = cy + r * math.sin(rad)
        x2 = cx - r * math.cos(rad)
        y2 = cy - r * math.sin(rad)
        self.e_canvas.coords(self.gear_id[1], x1, y1, x2, y2)
        
        # Spoke 2 (90 degrees offset)
        rad2 = rad + math.pi/2
        x3 = cx + r * math.cos(rad2)
        y3 = cy + r * math.sin(rad2)
        x4 = cx - r * math.cos(rad2)
        y4 = cy - r * math.sin(rad2)
        self.e_canvas.coords(self.gear_id[2], x3, y3, x4, y4)
