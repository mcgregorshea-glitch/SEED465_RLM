import time
import threading
import tkinter as tk
from tkinter import messagebox

class MotionControls:
    """
    Handles higher-level motion logic including jogging, homing, 
    coordinate centering, and collision testing.
    """
    def __init__(self, gui_instance) -> None:
        self.gui = gui_instance
        self.serial_engine = gui_instance.serial_engine
        
    def jog(self, axis: str, direction: int) -> None:
        """Sends a relative move command for manual jogging."""
        if not self.serial_engine.is_connected():
            messagebox.showerror("Error", "Not connected to a printer.")
            return

        try:
            if axis == 'E':
                step = float(self.gui.rotation_step_var.get())
                speed = float(self.gui.rotation_feedrate_var.get())
            else:
                step = float(self.gui.jog_step_var.get())
                speed = float(self.gui.jog_feedrate_var.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid step or speed value.")
            return

        dist = step * direction
        cmd = f"G91\nG1 {axis}{dist:.3f} F{speed:.0f}\nG90"
        self.gui._send_manual_command(cmd)

        # Optimistically update the UI state
        if axis == 'X' and self.gui.last_cmd_abs_x is not None: 
            self.gui.last_cmd_abs_x += dist; self.gui.target_abs_x = self.gui.last_cmd_abs_x
        elif axis == 'Y' and self.gui.last_cmd_abs_y is not None: 
            self.gui.last_cmd_abs_y += dist; self.gui.target_abs_y = self.gui.last_cmd_abs_y
        elif axis == 'Z' and self.gui.last_cmd_abs_z is not None: 
            self.gui.last_cmd_abs_z += dist; self.gui.target_abs_z = self.gui.last_cmd_abs_z
        elif (axis == 'ROT' or axis == 'E') and self.gui.last_cmd_abs_rot is not None: 
            self.gui.last_cmd_abs_rot += dist; self.gui.target_abs_rot = self.gui.last_cmd_abs_rot
        
        self.gui.root.after(50, self.gui._update_all_displays)

    def home_all(self) -> None:
        """Initiates the homing sequence in a background thread."""
        if not self.serial_engine.is_connected():
            messagebox.showerror("Error", "Not connected to a printer.")
            return
            
        if messagebox.askyesno("Confirm Homing", "Perform full X Y Z homing?\n\nEnsure the bed is clear!"):
            threading.Thread(target=self._homing_sequence_worker, daemon=True).start()

    def _homing_sequence_worker(self) -> None:
        """Background worker that issues a full X/Y/Z home."""
        self.gui.message_queue.put(("LOG", ("Starting full homing sequence (G28)...", "INFO")))
        self.gui.message_queue.put(("SET_STATUS", "busy"))

        try:
            # Full home of all axes. The manual-command worker appends M114 so the
            # displayed position refreshes once the printer reports homing complete.
            self.gui._send_manual_command("G28")
        except Exception as e:
            self.gui.message_queue.put(("LOG", (f"Homing failed: {e}", "ERROR")))
        finally:
            self.gui.message_queue.put(("SET_STATUS", "idle"))

    def go_to_position(self, x=None, y=None, z=None, r=None, speed=2000) -> None:
        """Sends an absolute move command to a specific coordinate."""
        if not self.serial_engine.is_connected():
            messagebox.showerror("Error", "Not connected to a printer.")
            return
            
        parts = ["G1"]
        if x is not None: parts.append(f"X{x:.3f}")
        if y is not None: parts.append(f"Y{y:.3f}")
        if z is not None: parts.append(f"Z{z:.3f}")
        if r is not None: parts.append(f"E{r:.3f}")
        parts.append(f"F{speed}")
        
        cmd = "G90\n" + " ".join(parts)
        self.gui._send_manual_command(cmd)
        
        # Optimistically update the UI state
        if x is not None: self.gui.last_cmd_abs_x = x; self.gui.target_abs_x = x
        if y is not None: self.gui.last_cmd_abs_y = y; self.gui.target_abs_y = y
        if z is not None: self.gui.last_cmd_abs_z = z; self.gui.target_abs_z = z
        if r is not None: self.gui.last_cmd_abs_rot = r; self.gui.target_abs_rot = r
        self.gui.root.after(50, self.gui._update_all_displays)

    def go_to_center(self) -> None:
        """Moves the printer to the defined center (0,0,0) coordinates."""
        try:
            cx = float(self.gui.center_x_var.get())
            cy = float(self.gui.center_y_var.get())
            cz = float(self.gui.center_z_var.get())
            crot = float(self.gui.center_rot_var.get())
            self.go_to_position(cx, cy, cz, crot)
        except ValueError:
            messagebox.showerror("Error", "Invalid center coordinates.")
