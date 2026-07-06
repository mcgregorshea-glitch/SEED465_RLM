import tkinter as tk
from tkinter import ttk
import sys
import os
import customtkinter as ctk

# Add relevant directories to sys.path to ensure modules like rlm_api are findable
script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(os.path.dirname(script_dir))
gui_src_dir = os.path.join(project_root, "legacy", "gui", "src")
seed_gui_dir = os.path.join(project_root, "legacy")

if gui_src_dir not in sys.path:
    sys.path.append(gui_src_dir)
if seed_gui_dir not in sys.path:
    sys.path.append(seed_gui_dir)
if script_dir not in sys.path:
    sys.path.append(script_dir)

import utils
from generator_panel import PatternGeneratorGUI
from sender_panel import GCodeSenderGUI
from vivigo_panel import VivigoPanel
from manual_movement_panel import ManualMovementPanel

class SEEDApplication:
    """
    The SEEDApplication class serves as the main entry point and root orchestrator for the SEED Control Center.
    
    It handles the initialization of the primary Tkinter window, manages navigation between different 
    application modules (Pattern Generator and G-Code Sender) via a tabbed interface, and implements 
    global safety features such as the Emergency Stop (E-STOP).
    
    To extend the system with new functionality, engineers should create new GUI panel classes 
    and register them as additional tabs within the 'notebook' container.
    """
    def __init__(self, root):
        """
        Initializes the main application environment, styles, and functional panels.
        
        Args:
            root (ctk.CTk): The root customtkinter window instance.
        """
        self.root = root # Primary Tkinter window instance used as the application's root
        self.root.title("⚡ SEED Control Center")
        self.root.geometry("1100x800")
        self.root.minsize(850, 650)
        self.root.configure(bg=utils.COLOR_BG)
        
        # Apply the global stylesheet for consistent visual theme
        utils.setup_global_styling(self.root)

        # Initialize shared EventBus for cross-panel communication
        self.event_bus = utils.EventBus()

        # Set customtkinter appearance
        ctk.set_appearance_mode("Dark")
        
        # Tab layout configuration
        self.notebook = ttk.Notebook(self.root, style='TNotebook') # Multi-tab container for navigating between system modules
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))
        
        # Create tab containers
        self.generator_tab = ttk.Frame(self.notebook, style='TFrame') # Logical container frame for the scan pattern design interface
        self.manual_tab = ttk.Frame(self.notebook, style='TFrame') # Logical container for manual profile creation
        self.sender_tab = ttk.Frame(self.notebook, style='TFrame') # Logical container frame for hardware communication and logging
        self.vivigo_tab = ttk.Frame(self.notebook, style='TFrame') # Logical container frame for Hub telemetry and control

        self.notebook.add(self.generator_tab, text="∿ Pattern Generator")
        self.notebook.add(self.manual_tab, text="✎ Manual Movement")
        self.notebook.add(self.sender_tab, text="❱ G-Code Sender")
        self.notebook.add(self.vivigo_tab, text="⚡ Vivigo Hub")
        
        # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        # CRITICAL ARCHITECTURAL CONSTRAINT:
        # The PoP Analysis panel MUST remain a SUBTAB inside VivigoPanel (vivigo_panel.py).
        # DO NOT add it as a fifth top-level tab to this notebook.
        # This 4-tab layout is an IMMUTABLE CONSTRAINT. See IMMUTABLE_CONSTRAINTS.md.
        # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

        # Instantiate the functional panels into their respective tabs
        self.generator_panel = PatternGeneratorGUI(self.generator_tab) # Logic and UI controller for scan path design and G-Code export
        self.manual_panel = ManualMovementPanel(self.manual_tab, controller=self) # Logic and UI controller for manual test sequencing
        self.manual_panel.pack(fill=tk.BOTH, expand=True)

        self.vivigo_panel = VivigoPanel(self.vivigo_tab, event_bus=self.event_bus) # Logic and UI controller for Hub communication
        self.vivigo_panel.pack(fill=tk.BOTH, expand=True)

        self.sender_panel = GCodeSenderGUI(self.sender_tab,
                                          vivigo_panel=self.vivigo_panel,
                                          pop_viz_panel=self.vivigo_panel.pop_panel,
                                          event_bus=self.event_bus)
        
        # Connect "Send to Sender" integration via callback
        self.generator_panel.on_send_to_sender = self._handle_send_to_sender
        self.manual_panel.on_send_to_sender = self._handle_send_to_sender
        
        # Initialize global safety UI
        self._setup_estop_button()
        
        self.root.bind_all('<Key>', self._handle_key_press)
        self.root.bind('<Escape>', self.trigger_estop)
        
    def _handle_key_press(self, e):
        # Ignore if the user is typing in any text or entry field
        if getattr(e.widget, 'winfo_class', lambda: '')() in ('Entry', 'TEntry', 'Text', 'TCombobox', 'TSpinbox'): return

        # Ensure we only process jog commands if the sender tab is currently visible
        if self.notebook.select() != str(self.sender_tab):
            return
        k = e.keysym.lower()
        if k == 'space': self.trigger_estop()
        elif k == 'w': self.sender_panel._jog('Y', -1)
        elif k == 's': self.sender_panel._jog('Y', 1)
        elif k == 'a': self.sender_panel._jog('X', -1)
        elif k == 'd': self.sender_panel._jog('X', 1)
        
    def _handle_send_to_sender(self, filepath):
        """
        Manages the integration workflow between the Pattern Generator and the G-Code Sender.
        
        This method is called when the user successfully generates a scan path and opts to 
        transfer it directly to the control interface. It loads the G-Code into the sender 
        and automatically switches focus to the Sender tab.

        Args:
            filepath (str): The absolute path to the generated G-Code file.
        """
        # 1. Load the G-Code into the sender panel
        self.sender_panel.load_gcode_file(filepath)
        # 2. Switch focus to the G-Code Sender tab (index 2)
        self.notebook.select(self.sender_tab)

    def trigger_estop(self, event=None):
        """Invokes the emergency stop routine in the sender panel and disconnects Hub."""
        if hasattr(self, 'sender_panel') and hasattr(self.sender_panel, 'emergency_stop'):
            self.sender_panel.emergency_stop()
        if hasattr(self, 'vivigo_panel') and hasattr(self.vivigo_panel, 'disconnect_hub'):
            self.vivigo_panel.disconnect_hub()

    def _setup_estop_button(self):
        """
        Configures and displays the global Emergency Stop (E-STOP) overlay.
        
        The E-STOP is implemented as a persistent canvas-based hexagonal button that remains 
        accessible regardless of which tab is active. Clicking this button immediately 
        invokes the emergency stop routine on the sender panel to halt all hardware operations.
        """
        # E-STOP Button (Hexagonal using Canvas)
        self.estop_canvas = tk.Canvas(self.root, width=80, height=80, bg=utils.COLOR_BG, highlightthickness=0) # Canvas element for rendering the global emergency stop button
        self.estop_canvas.place(relx=1.0, rely=0.0, x=-10, y=10, anchor="ne")
        
        # Draw a hexagon. Coordinates for an 80x80 hexagon centered on canvas
        hex_points = [40, 5, 75, 20, 75, 60, 40, 75, 5, 60, 5, 20]
        self.hex_id = self.estop_canvas.create_polygon(hex_points, fill=utils.COLOR_ACCENT_RED, outline=utils.COLOR_BORDER, width=2) # Unique identifier for the drawn hexagonal E-STOP shape
        self.text_id = self.estop_canvas.create_text(40, 40, text="ESTOP", font=("Consolas", 12, "bold"), fill=utils.COLOR_TEXT_PRIMARY) # Unique identifier for the 'ESTOP' text overlay on the canvas
        
        # Bind user clicks to the E-STOP logic
        self.estop_canvas.tag_bind(self.hex_id, '<Button-1>', self.trigger_estop)
        self.estop_canvas.tag_bind(self.text_id, '<Button-1>', self.trigger_estop)
        
        # Visual feedback: Hover effects for interactivity
        def on_enter(event):
            """Brighter color on mouse hover to indicate clickability."""
            self.estop_canvas.itemconfig(self.hex_id, fill="#ff1111")
            self.estop_canvas.config(cursor="hand2")
            
        def on_leave(event):
            """Restore original color when mouse leaves."""
            self.estop_canvas.itemconfig(self.hex_id, fill=utils.COLOR_ACCENT_RED)
            self.estop_canvas.config(cursor="")
            
        self.estop_canvas.tag_bind(self.hex_id, '<Enter>', on_enter)
        self.estop_canvas.tag_bind(self.text_id, '<Enter>', on_enter)
        self.estop_canvas.tag_bind(self.hex_id, '<Leave>', on_leave)
        self.estop_canvas.tag_bind(self.text_id, '<Leave>', on_leave)

def main():
    """
    Application entry point.
    
    Initializes the Tkinter event loop, creates the SEEDApplication instance, 
    and configures a clean shutdown protocol to ensure hardware-related threads 
    are terminated when the window is closed.
    """
    root = ctk.CTk()
    app = SEEDApplication(root)
    
    # Handle closing cleanly to terminate sender threads if needed
    def on_closing():
        """Clean shutdown handler."""
        if hasattr(app, 'sender_panel'):
            app.sender_panel.on_closing()
        else:
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    root.mainloop()

if __name__ == "__main__":
    main()
