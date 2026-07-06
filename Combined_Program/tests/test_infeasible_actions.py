import unittest
import tkinter as tk
from unittest.mock import MagicMock, patch
import os
import sys
import json

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src'))

from generator_panel import PatternGeneratorGUI

class TestInfeasibleActions(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.frame = tk.Frame(self.root)
        self.gui = PatternGeneratorGUI(self.frame)

    def tearDown(self):
        self.root.destroy()

    def test_infeasible_action_exclusion_from_gcode(self):
        """Test that infeasible actions are excluded from generated G-code."""
        # Setup a simple pattern
        params = {
            'x_min': 0, 'x_max': 10, 'x_step': 10,
            'y_min': 0, 'y_max': 0, 'y_step': 10,
            'z_min': 0, 'z_max': 0, 'z_step': 10,
            'rot_min': 0, 'rot_max': 0, 'rot_step': 10,
            'travelspeed': 3000, 'pause_time': 1
        }
        # Pattern will be [(0,0,0,0), (10,0,0,0)]
        pattern = [{'x': 0.0, 'y': 0.0, 'z': 0.0, 'rotation': 0.0}, 
                   {'x': 10.0, 'y': 0.0, 'z': 0.0, 'rotation': 0.0}]
        
        # Hub actions: one feasible, one infeasible
        self.gui.hub_actions = [
            {'x': tk.StringVar(value="0"), 'y': tk.StringVar(value="0"), 'z': tk.StringVar(value="0"), 
             'rot': tk.StringVar(value="0"), 'type': tk.StringVar(value="WPT Start"), 'color': "green"},
            {'x': tk.StringVar(value="5"), 'y': tk.StringVar(value="0"), 'z': tk.StringVar(value="0"), 
             'rot': tk.StringVar(value="0"), 'type': tk.StringVar(value="WPT Stop"), 'color': "red"}
        ]
        
        gcode_lines = list(self.gui.create_gcode(pattern, params, 2))
        
        # Verify "WPT Start" is in G-code (feasible)
        # Verify "WPT Stop" is NOT in G-code (infeasible)
        wpt_start_found = any("param=WPT Start value=true" in line for line in gcode_lines)
        wpt_stop_found = any("param=WPT Stop value=true" in line for line in gcode_lines)
        
        self.assertTrue(wpt_start_found, "Feasible action should be in G-code")
        self.assertFalse(wpt_stop_found, "Infeasible action should be excluded from G-code")

    def test_infeasible_action_flagging_in_table(self):
        """Test that infeasible actions are marked as such."""
        # Setup pattern: only point (0,0,0,0)
        self.gui._get_params_silently = MagicMock(return_value={
            'x_min': 0, 'x_max': 0, 'x_step': 10,
            'y_min': 0, 'y_max': 0, 'y_step': 10,
            'z_min': 0, 'z_max': 0, 'z_step': 10,
            'rot_min': 0, 'rot_max': 0, 'rot_step': 10,
            'travelspeed': 3000, 'pause_time': 1
        })
        
        # Hub actions: one feasible (0,0,0,0), one infeasible (10,10,10,10)
        self.gui.hub_actions = [
            {'x': tk.StringVar(value="0"), 'y': tk.StringVar(value="0"), 'z': tk.StringVar(value="0"), 
             'rot': tk.StringVar(value="0"), 'type': tk.StringVar(value="WPT Start"), 'color': "green"},
            {'x': tk.StringVar(value="10"), 'y': tk.StringVar(value="10"), 'z': tk.StringVar(value="10"), 
             'rot': tk.StringVar(value="10"), 'type': tk.StringVar(value="WPT Stop"), 'color': "red"}
        ]
        
        # Verify that we can identify which is feasible
        feasible_mask = self.gui._get_infeasible_actions_mask()
        self.assertFalse(feasible_mask[0], "First action should be feasible (False in mask)")
        self.assertTrue(feasible_mask[1], "Second action should be infeasible (True in mask)")

    @patch('generator_panel.messagebox.askyesno')
    def test_export_warning_on_infeasible_actions(self, mock_askyesno):
        """Test that a warning popup appears when exporting with infeasible actions."""
        # Setup infeasible action
        self.gui.hub_actions = [
            {'x': tk.StringVar(value="999"), 'y': tk.StringVar(value="0"), 'z': tk.StringVar(value="0"), 
             'rot': tk.StringVar(value="0"), 'type': tk.StringVar(value="WPT Stop"), 'color': "red"}
        ]
        
        # Mock dependencies for _start_generation_process
        self.gui._get_params_silently = MagicMock(return_value={
            'x_min': 0, 'x_max': 10, 'x_step': 10,
            'y_min': 0, 'y_max': 0, 'y_step': 10,
            'z_min': 0, 'z_max': 0, 'z_step': 10,
            'rot_min': 0, 'rot_max': 0, 'rot_step': 10,
            'travelspeed': 3000, 'pause_time': 1
        })
        self.gui.create_gcode = MagicMock(return_value=["G1 X0"])
        
        # If user says "No" to the warning, it should abort (not call create_gcode)
        mock_askyesno.return_value = False
        self.gui._start_generation_process()
        self.assertTrue(mock_askyesno.called, "Warning popup should be shown")
        self.assertFalse(self.gui.create_gcode.called, "Generation should abort if user says No")

if __name__ == '__main__':
    unittest.main()
