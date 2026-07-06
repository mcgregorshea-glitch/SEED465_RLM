import tkinter as tk
from tkinter import messagebox
import queue
import sys
import os
import time

# Add parent directory to path to allow imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from sim.serial_server import SerialServer
from sim.components.bed_viz import BedViz
from sim.components.state_panel import StatePanel
from sim.components.viz_3d import Viz3D
from sim.components.terminal_panel import TerminalPanel

class FakePrinterApp:
    """
    Main Application for the Fake Printer Emulator.
    """
    def __init__(self, root, port="COM9"):
        self.root = root
        self.root.title(f"Marlin Sim - Ender 3 Emulator ({port})")
        self.root.geometry("1100x850") # Increased height for terminal
        self.root.configure(bg="#1e1e1e")
        
        self.port = port
        
        # --- UI Components ---
        self.main_frame = tk.Frame(self.root, bg="#1e1e1e")
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # --- TOP SECTION (Visualizations & Controls) ---
        self.top_section = tk.Frame(self.main_frame, bg="#1e1e1e")
        self.top_section.pack(fill=tk.BOTH, expand=True)

        # Left: Visualizations
        self.viz_container = tk.Frame(self.top_section, bg="#1e1e1e")
        self.viz_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 2D Bed
        self.bed_viz = BedViz(self.viz_container)
        self.bed_viz.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        # 3D View
        self.viz_3d = Viz3D(self.viz_container)
        self.viz_3d.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=10)
        
        # Right: State Indicators & Controls
        self.ctrl_frame = tk.Frame(self.top_section, bg="#1e1e1e")
        self.ctrl_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=20)
        
        self.state_panel = StatePanel(self.ctrl_frame)
        self.state_panel.pack(fill=tk.X, pady=10)
        
        # Controls
        self.btn_frame = tk.LabelFrame(self.ctrl_frame, text="Controls", padx=10, pady=10)
        self.btn_frame.pack(fill=tk.X, pady=10)
        
        self.trail_var = tk.BooleanVar(value=True)
        self.trail_check = tk.Checkbutton(
            self.btn_frame, text="Show Path Trail", 
            variable=self.trail_var,
            command=self._on_trail_toggle
        )
        self.trail_check.pack(anchor=tk.W)
        
        # Link 2D trail var
        self.bed_viz.show_trail = self.trail_var
        
        tk.Button(self.btn_frame, text="Clear Trail", command=self._clear_all_trails).pack(fill=tk.X, pady=2)
        tk.Button(self.btn_frame, text="Reset / Home (G28)", command=self._manual_home).pack(fill=tk.X, pady=2)
        
        # Speed-up Factor
        self.speed_frame = tk.LabelFrame(self.ctrl_frame, text="Simulation Speed", padx=10, pady=10)
        self.speed_frame.pack(fill=tk.X, pady=10)
        
        self.speed_factor = tk.DoubleVar(value=1.0)
        self.speed_slider = tk.Scale(
            self.speed_frame, from_=1.0, to=500.0, 
            orient=tk.HORIZONTAL, variable=self.speed_factor,
            label="Multiplier (x)"
        )
        self.speed_slider.pack(fill=tk.X)

        # --- BOTTOM SECTION (Terminal) ---
        self.terminal = TerminalPanel(self.main_frame)
        self.terminal.config(height=200)
        self.terminal.pack(fill=tk.BOTH, expand=False, pady=(10, 0))
        
        self.status_label = tk.Label(self.root, text=f"Status: Waiting for {port}...", fg="orange", bg="#1e1e1e")
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)
        
        # --- Serial Server ---
        self.server = SerialServer(port=self.port)
        self.server.start()
        
        # Start update loop (30 FPS)
        self.root.after(33, self._update_loop)

    def _update_loop(self):
        """Pulls updates from the serial thread and refreshes the GUI."""
        # Sync speed factor to server thread
        self.server.speed_factor = self.speed_factor.get()
        
        # 1. Drain logs (limit to 20 per frame to avoid GUI hang)
        count = 0
        try:
            while count < 20:
                direction, msg = self.server.log_queue.get_nowait()
                self.terminal.log_command(direction, msg)
                count += 1
        except queue.Empty:
            pass

        # 2. Update GUI from current simulation state
        state = self.server.state.get_state()
        self.bed_viz.update_toolhead(state['X'], state['Y'])
        self.viz_3d.update_toolhead(state['X'], state['Y'], state['Z'])
        self.state_panel.update_state(state['Z'], state['E'])
        
        # 3. Check for serial connectivity status
        status_text = f"X:{state['X']:.1f} Y:{state['Y']:.1f} Z:{state['Z']:.1f}"
        if self.server.running:
             self.status_label.config(text=f"Status: Listening on {self.port} | {status_text}", fg="green")
        else:
             self.status_label.config(text=f"Status: Disconnected | {status_text}", fg="red")
        
        # Re-schedule (33ms = 30 FPS - smoother for slow computers)
        self.root.after(33, self._update_loop)

    def _on_trail_toggle(self):
        visible = self.trail_var.get()
        self.bed_viz.set_trail_visibility(visible)
        self.viz_3d.set_trail_visibility(visible)

    def _clear_all_trails(self):
        self.bed_viz.clear_trail()
        self.viz_3d.clear_trail()

    def _manual_home(self):
        self.server.state.home()
        self.bed_viz.update_toolhead(0, 0)
        self.viz_3d.update_toolhead(0, 0, 0)
        self.state_panel.update_state(0, 0)
        self._clear_all_trails()

    def on_closing(self):
        self.server.stop()
        self.root.destroy()

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Marlin Fake Printer Emulator")
    parser.add_argument("--port", type=str, default="COM9", help="Serial port to listen on")
    args = parser.parse_args()
    
    root = tk.Tk()
    app = FakePrinterApp(root, port=args.port)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
