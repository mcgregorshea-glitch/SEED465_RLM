import tkinter as tk
from tkinter import ttk
import utils
from generator_panel import PatternGeneratorGUI
from sender_panel import GCodeSenderGUI

class SEEDApplication:
    def __init__(self, root):
        self.root = root
        self.root.title("⚡ SEED Control Center")
        self.root.geometry("1100x800")
        self.root.minsize(850, 650)
        self.root.configure(bg=utils.COLOR_BG)
        
        # Apply the global stylesheet
        utils.setup_global_styling(self.root)
        
        # Tab layout
        self.notebook = ttk.Notebook(self.root, style='TNotebook')
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))
        
        # Create tabs
        self.generator_tab = ttk.Frame(self.notebook, style='TFrame')
        self.sender_tab = ttk.Frame(self.notebook, style='TFrame')
        
        self.notebook.add(self.generator_tab, text="∿ Pattern Generator")
        self.notebook.add(self.sender_tab, text="❱ G-Code Sender")
        
        # Instantiate the panels into their respective tabs
        self.generator_panel = PatternGeneratorGUI(self.generator_tab)
        self.sender_panel = GCodeSenderGUI(self.sender_tab)
        
        # Connect "Send to Sender" integration
        self.generator_panel.on_send_to_sender = self._handle_send_to_sender
        
        self._setup_estop_button()
        
    def _handle_send_to_sender(self, filepath):
        """Called when a user generates G-code and wants to send it to the Sender."""
        # 1. Load the G-Code securely
        self.sender_panel.load_gcode_file(filepath)
        # 2. Switch to the G-Code Sender tab (index 1)
        self.notebook.select(self.sender_tab)

    def _setup_estop_button(self):
        # E-STOP Button (Hexagonal using Canvas)
        # We overlay this on the root window so it's always accessible
        self.estop_canvas = tk.Canvas(self.root, width=80, height=80, bg=utils.COLOR_BG, highlightthickness=0)
        self.estop_canvas.place(relx=1.0, rely=0.0, x=-10, y=10, anchor="ne")
        
        # Draw a hexagon. Coordinates for an 80x80 hexagon:
        # (40, 0), (75, 20), (75, 60), (40, 80), (5, 60), (5, 20)
        hex_points = [40, 5, 75, 20, 75, 60, 40, 75, 5, 60, 5, 20]
        self.hex_id = self.estop_canvas.create_polygon(hex_points, fill=utils.COLOR_ACCENT_RED, outline=utils.COLOR_BORDER, width=2)
        self.text_id = self.estop_canvas.create_text(40, 40, text="ESTOP", font=("Consolas", 12, "bold"), fill=utils.COLOR_TEXT_PRIMARY)
        
        # Bind clicks to emergency stop
        def trigger_estop(event):
            if hasattr(self, 'sender_panel') and hasattr(self.sender_panel, 'emergency_stop'):
                self.sender_panel.emergency_stop()
                
        self.estop_canvas.tag_bind(self.hex_id, '<Button-1>', trigger_estop)
        self.estop_canvas.tag_bind(self.text_id, '<Button-1>', trigger_estop)
        
        # Make hover effect
        def on_enter(event):
            self.estop_canvas.itemconfig(self.hex_id, fill="#ff1111")
            self.estop_canvas.config(cursor="hand2")
            
        def on_leave(event):
            self.estop_canvas.itemconfig(self.hex_id, fill=utils.COLOR_ACCENT_RED)
            self.estop_canvas.config(cursor="")
            
        self.estop_canvas.tag_bind(self.hex_id, '<Enter>', on_enter)
        self.estop_canvas.tag_bind(self.text_id, '<Enter>', on_enter)
        self.estop_canvas.tag_bind(self.hex_id, '<Leave>', on_leave)
        self.estop_canvas.tag_bind(self.text_id, '<Leave>', on_leave)

def main():
    root = tk.Tk()
    app = SEEDApplication(root)
    
    # Handle closing cleanly to terminate sender threads if needed
    def on_closing():
        if hasattr(app.sender_panel, 'is_sending') and app.sender_panel.is_sending:
            import time
            time.sleep(0.5)
        root.destroy()
        
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()
