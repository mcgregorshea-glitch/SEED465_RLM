import unittest
import tkinter as tk
from unittest.mock import MagicMock
import os
import sys
import json

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src'))

from generator_panel import PatternGeneratorGUI

class TestGeneratorIO(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.frame = tk.Frame(self.root)
        self.gui = PatternGeneratorGUI(self.frame)

    def tearDown(self):
        self.root.destroy()

    def test_gcode_header_generation(self):
        """Test that create_gcode generates a JSON header."""
        # Set GUI values directly
        self.gui.x_symmetric.set(False)
        self.gui.x_min.delete(0, tk.END)
        self.gui.x_min.insert(0, "-10")
        self.gui.x_max.delete(0, tk.END)
        self.gui.x_max.insert(0, "10")
        
        params = {
            'x_min': -10, 'x_max': 10, 'x_step': 5,
            'y_min': -10, 'y_max': 10, 'y_step': 5,
            'z_min': 0, 'z_max': 5, 'z_step': 5,
            'rot_min': 0, 'rot_max': 0, 'rot_step': 5,
            'travelspeed': 3000, 'pause_time': 1
        }
        # Mock hub_actions
        self.gui.hub_actions = [
            {'x': tk.StringVar(value="5"), 'y': tk.StringVar(value="5"), 'z': tk.StringVar(value="0"), 
             'rot': tk.StringVar(value="0"), 'type': tk.StringVar(value="WPT Start"), 'color': "green"}
        ]
        
        pattern = [{'x': 0, 'y': 0, 'z': 0, 'rotation': 0}]
        gcode_lines = list(self.gui.create_gcode(pattern, params, 1))
        
        # Check if any line contains the JSON header
        header_line = None
        for line in gcode_lines:
            if line.startswith("; PARAMS:"):
                header_line = line
                break
        
        self.assertIsNotNone(header_line, "G-code should contain a ; PARAMS: header")
        header_json = header_line.replace("; PARAMS:", "").strip()
        data = json.loads(header_json)
        self.assertEqual(data['x_min'], "-10")
        self.assertEqual(len(data['hub_actions']), 1)
        self.assertEqual(data['hub_actions'][0]['type'], "WPT Start")

    def test_load_parameters_from_gcode(self):
        """Test that parameters are correctly loaded from a G-code file."""
        test_file = "test_load.gcode"
        header_data = {
            "profile_name": "LOADED_TEST",
            "x_symmetric": False,
            "x_min": -25.0, "x_max": 25.0, "x_step": 2.0,
            "y_symmetric": True, "y_offset": 30.0, "y_step": 1.0,
            "z_symmetric": False, "z_min": 10.0, "z_max": 20.0, "z_step": 10.0,
            "rot_symmetric": True, "rot_offset": 0.0, "rot_step": 5.0,
            "travelspeed": 1500, "pause_time": 2.5,
            "hub_actions": [
                {"x": "1.23", "y": "4.56", "z": "7.89", "rot": "0", "type": "WPT Stop"}
            ]
        }
        with open(test_file, 'w') as f:
            f.write("; SEED Pattern\n")
            f.write(f"; PARAMS: {json.dumps(header_data)}\n")
            f.write("G1 X0 Y0 Z0\n")

        try:
            # Implement this method
            self.gui._load_parameters_from_gcode_file(test_file)
            
            # Verify basic fields
            self.assertEqual(self.gui.profile_name_var.get(), "LOADED_TEST")
            self.assertEqual(self.gui.travelspeed.get(), "1500")
            self.assertEqual(self.gui.pause_time.get(), "2.5")
            
            # Verify symmetry
            self.assertFalse(self.gui.x_symmetric.get())
            self.assertTrue(self.gui.y_symmetric.get())
            
            # Verify coordinates
            self.assertEqual(self.gui.x_min.get(), "-25.0")
            self.assertEqual(self.gui.y_offset.get(), "30.0")
            
            # Verify Hub Actions
            self.assertEqual(len(self.gui.hub_actions), 1)
            self.assertEqual(self.gui.hub_actions[0]['type'].get(), "WPT Stop")
            self.assertEqual(self.gui.hub_actions[0]['x'].get(), "1.23")
            
        finally:
            if os.path.exists(test_file):
                os.remove(test_file)

if __name__ == '__main__':
    unittest.main()
