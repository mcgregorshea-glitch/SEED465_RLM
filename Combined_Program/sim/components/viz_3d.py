import tkinter as tk
import math
from typing import Tuple

class Viz3D(tk.Frame):
    """
    A lightweight 3D visualization using isometric projection on a Tkinter Canvas.
    No external dependencies (Matplotlib/OpenGL).
    """
    def __init__(self, master, bed_size=(220, 220, 250)):
        super().__init__(master, bg="#1e1e1e")
        self.dim_x, self.dim_y, self.dim_z = bed_size
        
        # Initial projection state
        self.scale = 0.5
        self.offset_x = 150
        self.offset_y = 250
        self.angle_x = math.radians(30)
        self.angle_y = math.radians(30)
        
        self.canvas = tk.Canvas(self, bg="#2d2d2d", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.toolhead = self.canvas.create_oval(0, 0, 0, 0, fill="red", outline="white")
        
        # Trail state
        self.show_trail = True
        self.last_pos = (0.0, 0.0, 0.0)
        self.trail_data = []
        self.trail_lines = []
        
        self.canvas.bind("<Configure>", self._on_resize)
        self.update_toolhead(0, 0, 0)

    def _project(self, x: float, y: float, z: float) -> Tuple[float, float]:
        """
        Standard Isometric projection:
        - View from Front-Center/Top
        - Origin (0,0,0) at the Bottom-Center-Back of the diamond
        """
        px = (x - y) * math.cos(self.angle_x) * self.scale
        # py: (x+y) moves DOWN (forward), -z moves UP (height)
        py = ((x + y) * math.sin(self.angle_y) - z) * self.scale
        return self.offset_x + px, self.offset_y + py

    def _draw_bounding_box(self):
        """Draws the wireframe cube and axis labels."""
        points = [
            (0, 0, 0), (self.dim_x, 0, 0), (self.dim_x, self.dim_y, 0), (0, self.dim_y, 0),
            (0, 0, self.dim_z), (self.dim_x, 0, self.dim_z), (self.dim_x, self.dim_y, self.dim_z), (0, self.dim_y, self.dim_z)
        ]
        proj = [self._project(*p) for p in points]
        edges = [
            (0,1), (1,2), (2,3), (3,0), # Bottom
            (4,5), (5,6), (6,7), (7,4), # Top
            (0,4), (1,5), (2,6), (3,7)  # Pillars
        ]
        for start, end in edges:
            self.canvas.create_line(proj[start], proj[end], fill="#555555", width=1)

        # Axis Helpers
        ox, oy = proj[0]
        xx, xy = proj[1]
        yx, yy = proj[3]
        zx, zy = proj[4]
        self.canvas.create_text(xx+10, xy, text="X+", fill="red", font=("Arial", 8))
        self.canvas.create_text(yx-10, yy, text="Y+", fill="green", font=("Arial", 8))
        self.canvas.create_text(zx, zy-10, text="Z+", fill="blue", font=("Arial", 8))

    def _on_resize(self, event):
        """Robust centering logic based on visual volume spread."""
        w, h = event.width, event.height
        
        # Total screen space required for the 220x220x250 volume
        total_w = (self.dim_x + self.dim_y) * math.cos(self.angle_x)
        total_h = self.dim_z + (self.dim_x + self.dim_y) * math.sin(self.angle_y)
        
        # Scale to fit with a 25% margin
        self.scale = min(w / total_w, h / total_h) * 0.75
        
        # Calculate visual midpoint of the volume
        vx_center = (self.dim_x - self.dim_y) / 2 * math.cos(self.angle_x) * self.scale
        vy_center = ((self.dim_x + self.dim_y) / 2 * math.sin(self.angle_y) - self.dim_z / 2) * self.scale
        
        self.offset_x = (w / 2) - vx_center
        self.offset_y = (h / 2) - vy_center
        
        self.redraw()

    def redraw(self):
        """Redraws bounding box, toolhead, and trails."""
        self.canvas.delete("all")
        self._draw_bounding_box()
        
        # Redraw toolhead
        lx, ly, lz = self.last_pos
        px, py = self._project(lx, ly, lz)
        r = 4
        self.toolhead = self.canvas.create_oval(px-r, py-r, px+r, py+r, fill="red", outline="white")
        
        # Redraw trails
        self.trail_lines = []
        if self.show_trail:
            for i in range(1, len(self.trail_data)):
                p1, p2 = self.trail_data[i-1], self.trail_data[i]
                x1, y1 = self._project(*p1)
                x2, y2 = self._project(*p2)
                self.trail_lines.append(self.canvas.create_line(x1, y1, x2, y2, fill="cyan", dash=(2, 2)))
        
        self.canvas.tag_raise(self.toolhead)

    def update_toolhead(self, x: float, y: float, z: float):
        """Moves toolhead and records significant movement to trail data."""
        # Only append if movement is significant
        dist = math.sqrt((x - self.last_pos[0])**2 + (y - self.last_pos[1])**2 + (z - self.last_pos[2])**2)
        if dist > 0.1:
            self.trail_data.append((x, y, z))
            
            px, py = self._project(x, y, z)
            if self.show_trail and len(self.trail_data) > 1:
                lx, ly, lz = self.trail_data[-2]
                lpx, lpy = self._project(lx, ly, lz)
                self.trail_lines.append(self.canvas.create_line(lpx, lpy, px, py, fill="cyan", dash=(2, 2)))
        
        self.last_pos = (x, y, z)
        px, py = self._project(x, y, z)
        r = 4
        self.canvas.coords(self.toolhead, px-r, py-r, px+r, py+r)
        self.canvas.tag_raise(self.toolhead)

    def clear_trail(self):
        self.trail_data = [self.last_pos]
        self.redraw()

    def set_trail_visibility(self, visible: bool):
        self.show_trail = visible
        self.redraw()
