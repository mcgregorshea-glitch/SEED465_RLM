import tkinter as tk
from tkinter import scrolledtext

class TerminalPanel(tk.Frame):
    """
    A high-fidelity terminal readout for displaying incoming serial commands.
    Features monospaced typography and high-contrast diagnostic colors.
    """
    def __init__(self, master):
        super().__init__(master, bg="#1e1e1e")
        
        self.label_frame = tk.LabelFrame(
            self, text="Serial Command Log", 
            fg="#cccccc", bg="#1e1e1e", 
            font=("Segoe UI", 9, "bold"),
            padx=5, pady=5
        )
        self.label_frame.pack(fill=tk.BOTH, expand=True)

        # ScrolledText widget for automatic scrollbar handling
        self.text_area = scrolledtext.ScrolledText(
            self.label_frame, 
            bg="#121212", 
            fg="#00ff41", # Classic Terminal Green
            insertbackground="white", 
            font=("Consolas", 10),
            padx=10, pady=10,
            state='disabled', # Start read-only
            borderwidth=0,
            highlightthickness=1,
            highlightbackground="#333333"
        )
        self.text_area.pack(fill=tk.BOTH, expand=True)

    def log_command(self, direction: str, message: str):
        """
        Appends a message to the terminal with direction indicators.
        'RX' for incoming from host, 'TX' for outgoing from printer.
        """
        self.text_area.configure(state='normal')
        
        indicator = ">> " if direction == "RX" else "<< "
        color = "#00ff41" if direction == "RX" else "#888888"
        
        # Insert at the end
        self.text_area.insert(tk.END, f"{indicator}{message}\n")
        
        # Keep only last 200 lines to maintain performance
        line_count = int(self.text_area.index('end-1c').split('.')[0])
        if line_count > 200:
            self.text_area.delete('1.0', '2.0')

        self.text_area.see(tk.END) # Auto-scroll
        self.text_area.configure(state='disabled')
