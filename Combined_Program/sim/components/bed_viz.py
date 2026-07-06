import tkinter as tk
import math
from typing import Dict, Tuple

class BedViz(tk.Frame):
    """
    A 2D visualization of the printer bed (XY plane).
    Displays the toolhead as a red dot and an optional path trail.
    """
    def __init__(self, master, bed_size: Tuple[float, float] = (220.0, 220.0)):
        super().__init__(master, bg="#1e1e1e")
        self.bed_w, self.bed_h = bed_size
        self.scale = 1.0
        
        self.canvas = tk.Canvas(self, bg="#2d2d2d", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Toolhead representation
        self.toolhead = self.canvas.create_oval(0, 0, 0, 0, fill="red", outline="white")
        
        # Trail state
        self.show_trail = tk.BooleanVar(value=True)
        self.last_pos: Tuple[float, float] = (0.0, 0.0)
        self.trail_data = [] # Store (x, y) coordinates instead of line IDs
        self.trail_lines = []
        
        # Initial dimensions (will be updated on resize)
        self.canvas_w = 400
        self.canvas_h = 400
        self.off_x = 20
        self.off_y = 20
        self.scale = 1.5 # Reasonable starting scale for 400x400
        
        self.canvas.bind("<Configure>", self._on_resize)
        self.update_toolhead(0.0, 0.0)
        self.redraw()

    def _on_resize(self, event):
        """Recalculate scale (1:1 aspect ratio) and center the bed."""
        # Padding
        pad = 20
        w = event.width - (pad * 2)
        h = event.height - (pad * 2)
        
        # Force 1:1 Aspect Ratio
        self.scale = min(w / self.bed_w, h / self.bed_h)
        
        # Calculate offsets to center the bed
        self.off_x = (event.width - (self.bed_w * self.scale)) / 2
        self.off_y = (event.height - (self.bed_h * self.scale)) / 2
        
        self.canvas_w = event.width
        self.canvas_h = event.height
        
        self.redraw()

    def redraw(self):
        """Redraws everything with centering offsets."""
        self.canvas.delete("all")
        
        # Draw bed boundary for clarity
        x1, y1 = self.off_x, self.off_y
        x2, y2 = x1 + (self.bed_w * self.scale), y1 + (self.bed_h * self.scale)
        self.canvas.create_rectangle(x1, y1, x2, y2, outline="#444444", dash=(2,2))
        
        # Redraw toolhead
        lx, ly = self.last_pos
        px = lx * self.scale + self.off_x
        py = self.canvas_h - (ly * self.scale) - self.off_y
        r = 5
        self.toolhead = self.canvas.create_oval(px-r, py-r, px+r, py+r, fill="red", outline="white")
        
        # Redraw trail
        self.trail_lines = []
        if self.show_trail.get():
            for i in range(1, len(self.trail_data)):
                p1, p2 = self.trail_data[i-1], self.trail_data[i]
                tx1 = p1[0] * self.scale + self.off_x
                ty1 = self.canvas_h - (p1[1] * self.scale) - self.off_y
                tx2 = p2[0] * self.scale + self.off_x
                ty2 = self.canvas_h - (p2[1] * self.scale) - self.off_y
                self.trail_lines.append(self.canvas.create_line(tx1, ty1, tx2, ty2, fill="cyan", dash=(2, 2)))

    def update_toolhead(self, x: float, y: float):
        """Updates position and records trail data with centering."""
        # Only append if movement is significant (e.g., > 0.1mm)
        dist = math.sqrt((x - self.last_pos[0])**2 + (y - self.last_pos[1])**2)
        if dist > 0.1:
            self.trail_data.append((x, y))
            
            # Draw new line segment
            if self.show_trail.get() and len(self.trail_data) > 1:
                p1 = self.trail_data[-2]
                tx1 = p1[0] * self.scale + self.off_x
                ty1 = self.canvas_h - (p1[1] * self.scale) - self.off_y
                px = x * self.scale + self.off_x
                py = self.canvas_h - (y * self.scale) - self.off_y
                self.trail_lines.append(self.canvas.create_line(tx1, ty1, px, py, fill="cyan", dash=(2, 2)))
        
        self.last_pos = (x, y)
        
        px = x * self.scale + self.off_x
        py = self.canvas_h - (y * self.scale) - self.off_y
        r = 5
        self.canvas.coords(self.toolhead, px-r, py-r, px+r, py+r)
        self.canvas.tag_raise(self.toolhead)

    def clear_trail(self):
        """Removes trail history."""
        self.trail_data = [self.last_pos]
        self.redraw()

    def set_trail_visibility(self, visible: bool):
        """Toggles trail visibility."""
        self.redraw() # Simplest way to toggle with persistent data
