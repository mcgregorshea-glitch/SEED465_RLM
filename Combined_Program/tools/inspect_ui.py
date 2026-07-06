import tkinter as tk
from tkinter import ttk

def print_hierarchy(widget, indent=0):
    """
    Recursively prints the widget hierarchy with layout info.
    Useful for debugging 'thick' tables and horizontal clipping.
    """
    try:
        # Get basic info
        name = widget.winfo_name()
        w_class = widget.winfo_class()
        w = widget.winfo_width()
        h = widget.winfo_height()
        
        # Get grid/pack info
        manager = widget.winfo_manager()
        layout = ""
        if manager == "grid":
            info = widget.grid_info()
            layout = f" | Grid(r={info.get('row')}, c={info.get('column')}, span={info.get('columnspan')})"
        elif manager == "pack":
            info = widget.pack_info()
            layout = f" | Pack(side={info.get('side')}, fill={info.get('fill')})"
            
        print("  " * indent + f"[{w_class}] {name} ({w}x{h}){layout}")
        
        for child in widget.winfo_children():
            print_hierarchy(child, indent + 1)
    except Exception as e:
        print("  " * indent + f"!! Error inspecting {widget}: {e}")

if __name__ == "__main__":
    print("UI Inspector Utility Loaded.")
    print("Usage: Call print_hierarchy(root_widget) during runtime.")
