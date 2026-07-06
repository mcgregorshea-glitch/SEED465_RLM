import tkinter as tk
from tkinter import ttk
import os
import sys
from pathlib import Path
import utils

# Add the paths to ensure imports work
script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(os.path.dirname(script_dir))
gui_src_dir = os.path.join(project_root, "legacy", "gui", "src")
seed_gui_dir = os.path.join(project_root, "legacy")

if gui_src_dir not in sys.path:
    sys.path.append(gui_src_dir)
if seed_gui_dir not in sys.path:
    sys.path.append(seed_gui_dir)

# Import the original GUI and supporting classes
from gui import GUI
from hierarchy import StartupArgs
import rlm_proj_vivigo
from pop_visualization_panel import POPVisualizationPanel

class VivigoPanel(ttk.Frame):
    """
    A GUI panel for controlling and monitoring the Vivigo Hub.
    This version includes subtabs for the original Vivigo GUI and POP Visualization.
    """
    def __init__(self, parent_frame, event_bus=None):
        super().__init__(parent_frame, style='Dark.TFrame')
        self.parent = parent_frame
        self.event_bus = event_bus
        
        # Internal Notebook for subtabs
        self.notebook = ttk.Notebook(self, style='TNotebook')
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Subtab 1: Vivigo GUI
        self.gui_tab = ttk.Frame(self.notebook, style='Dark.TFrame')
        self.notebook.add(self.gui_tab, text=" ⚡ VIVIGO CONTROLS ")
        
        # Load the project configuration for Vivigo Hub
        project = rlm_proj_vivigo.gen_project()
        startup_args = StartupArgs(
            exepath=Path(sys.argv[0]),
            project=project,
            comport=os.environ.get("SEED_SIM_HUB_PORT"),
            verbose=False,
            random=False,
            filelogging=True
        )
        
        self.gui = GUI(self.gui_tab, startup_args, embedded=True)
        self.gui.pack(fill=tk.BOTH, expand=True)

        # Subtab 2: PoP Visualization
        self.pop_tab = ttk.Frame(self.notebook, style='Dark.TFrame')
        self.notebook.add(self.pop_tab, text=" 📊 POP Analysis ")
        self.pop_panel = POPVisualizationPanel(self.pop_tab, controller=self)
        self.pop_panel.pack(fill=tk.BOTH, expand=True)

        # Provide compatibility with the SEED Control Center's existing integration
        self.rlm = self.gui

    def disconnect_hub(self):
        """
        Stops the communication backend and cleans up.
        Called by the main application during shutdown or emergency stop.
        """
        if hasattr(self.gui, 'backend') and self.gui.backend:
            self.gui.backend.end()
            
    def connect_hub(self, config_path, comport):
        """
        Compatibility method. In the embedded GUI, connection is managed 
        via the GUI's own port selection dropdown.
        """
        if hasattr(self.gui, 'port_selected'):
            self.gui.port_selected(comport)
            return True
        return False

    def get_active_port(self):
        """Returns the currently selected COM port from the internal GUI."""
        if hasattr(self.gui, 'port_select'):
            return self.gui.port_select.get()
        return None

if __name__ == "__main__":
    # Quick standalone test for development
    import customtkinter as ctk
    root = ctk.CTk()
    root.title("Vivigo Panel Standalone Test")
    root.geometry("1200x800")
    app = VivigoPanel(root)
    root.mainloop()
