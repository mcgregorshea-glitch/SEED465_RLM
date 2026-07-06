import unittest
import tkinter as tk
import os
import sys
import json

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src'))

from generator_panel import PatternGeneratorGUI

class TestGCodeInjection(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.frame = tk.Frame(self.root)
        self.gui = PatternGeneratorGUI(self.frame)

    def tearDown(self):
        self.root.destroy()

    def test_inline_command_injection(self):
        """Test that HUB_CMD is injected at the correct coordinate."""
        params = {
            'x_min': 0, 'x_max': 10, 'x_step': 10,
            'y_min': 0, 'y_max': 0, 'y_step': 1,
            'z_min': 0, 'z_max': 0, 'z_step': 1,
            'rot_min': 0, 'rot_max': 0, 'rot_step': 1,
            'travelspeed': 3000, 'pause_time': 1
        }
        
        # Set a Hub Action at (10, 0, 0, 0)
        self.gui.hub_actions = [
            {
                'x': tk.StringVar(value="10.0"), 
                'y': tk.StringVar(value="0.0"), 
                'z': tk.StringVar(value="0.0"), 
                'rot': tk.StringVar(value="0.0"), 
                'type': tk.StringVar(value="WPT Start"), 
                'color': "green"
            }
        ]
        
        pattern = [
            {'x': 0.0, 'y': 0.0, 'z': 0.0, 'rotation': 0.0},
            {'x': 10.0, 'y': 0.0, 'z': 0.0, 'rotation': 0.0}
        ]
        
        gcode_lines = list(self.gui.create_gcode(pattern, params, 2))
        
        # Look for the line moving to X10 and check if HUB_CMD follows
        found_move = False
        found_cmd = False
        for i, line in enumerate(gcode_lines):
            if "X10.000" in line and "G1" in line:
                found_move = True
                # Check next line
                if i + 1 < len(gcode_lines) and "; HUB_CMD param=WPT Start value=true" in gcode_lines[i+1]:
                    found_cmd = True
                    break
        
        self.assertTrue(found_move, "Should have found the move to X10")
        self.assertTrue(found_cmd, "Should have found the HUB_CMD immediately after the move")

if __name__ == '__main__':
    unittest.main()
