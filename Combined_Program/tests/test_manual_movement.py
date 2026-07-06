import unittest
import tkinter as tk
from tkinter import ttk
from unittest.mock import MagicMock, patch
import sys
import os
import re

# Add src to path so 'import utils' works
script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
src_dir = os.path.join(project_root, "src")
if src_dir not in sys.path:
    sys.path.append(src_dir)

from manual_movement_panel import ManualMovementPanel

class TestManualMovementPanel(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        # Mock controller
        class MockController:
            def _handle_send_to_sender(self, path):
                self.last_sent_path = path
        self.controller = MockController()
        self.panel = ManualMovementPanel(self.root, controller=self.controller)

    def tearDown(self):
        self.root.destroy()

    def test_add_action(self):
        initial_count = len(self.panel.actions)
        self.panel._add_action_row()
        self.assertEqual(len(self.panel.actions), initial_count + 1)
        self.assertEqual(self.panel.actions[-1]['type'].get(), "Move")

    def test_remove_action(self):
        self.panel._add_action_row()
        count_after_add = len(self.panel.actions)
        self.panel._remove_action_row(0)
        self.assertEqual(len(self.panel.actions), count_after_add - 1)

    def test_speed_inheritance(self):
        # Add first move with specific speed
        self.panel._add_action_row()
        self.panel.actions[0]['type'].set("Move")
        self.panel.actions[0]['speed'].set("5000")
        
        # Add second row (not move)
        self.panel._add_action_row()
        self.panel.actions[1]['type'].set("Pause")
        
        # Add third row (move) - should inherit from first
        self.panel._add_action_row()
        self.assertEqual(self.panel.actions[2]['speed'].get(), "5000")

    def test_gcode_export(self):
        self.panel._add_action_row()
        self.panel.actions[0]['x'].set("10")
        self.panel.actions[0]['y'].set("20")
        
        test_file = "test_manual_export.gcode"
        try:
            self.panel._write_gcode_file(test_file)
            self.assertTrue(os.path.exists(test_file))
            with open(test_file, 'r') as f:
                content = f.read()
                self.assertIn("G1 X10 Y20", content)
        finally:
            if os.path.exists(test_file):
                os.remove(test_file)

    @patch('manual_movement_panel.messagebox.showinfo')
    @patch('manual_movement_panel.messagebox.showwarning')
    @patch('manual_movement_panel.messagebox.showerror')
    def test_gcode_import(self, mock_error, mock_warning, mock_info):
        # Create a mock G-code file with mixed actions
        test_gcode = (
            "G1 X10 Y10 Z5 E0 F3000\n"
            "; HUB_CMD param=WPT Start value=true\n"
            "; WAIT seconds=2.5\n"
            "G1 X20 Y20 Z10 E45 F5000\n"
            "; WAIT seconds=1.0\n"
        )
        test_file = "test_manual_import.gcode"
        try:
            with open(test_file, 'w') as f:
                f.write(test_gcode)
            
            # Mock filedialog.askopenfilename to return our test file
            import tkinter.filedialog as filedialog
            original_ask = filedialog.askopenfilename
            filedialog.askopenfilename = lambda **kwargs: test_file
            
            try:
                self.panel._load_gcode()
                
                # Verify recovery
                self.assertEqual(len(self.panel.actions), 5)
                self.assertEqual(self.panel.actions[0]['type'].get(), "Move")
                self.assertEqual(self.panel.actions[1]['type'].get(), "Hub Action")
                self.assertEqual(self.panel.actions[1]['sub_type'].get(), "WPT Start")
                self.assertEqual(self.panel.actions[2]['type'].get(), "Pause")
                self.assertEqual(self.panel.actions[2]['duration'].get(), "2500")
                self.assertEqual(self.panel.actions[3]['x'].get(), "20")
                self.assertEqual(self.panel.actions[3]['rot'].get(), "45")
                self.assertEqual(self.panel.actions[4]['type'].get(), "Pause")
                self.assertEqual(self.panel.actions[4]['duration'].get(), "1000")
                
                self.assertTrue(mock_info.called)
                
            finally:
                filedialog.askopenfilename = original_ask
                
        finally:
            if os.path.exists(test_file):
                os.remove(test_file)

    def test_reordering(self):
        # Add two moves
        self.panel._add_action_row() # idx 0
        self.panel.actions[0]['x'].set("10")
        self.panel._add_action_row() # idx 1
        self.panel.actions[1]['x'].set("20")
        
        # Move first one down
        self.panel._move_action(0, 1)
        self.assertEqual(self.panel.actions[0]['x'].get(), "20")
        self.assertEqual(self.panel.actions[1]['x'].get(), "10")

    def test_mid_sequence_insertion(self):
        # Add two moves
        self.panel._add_action_row() # idx 0
        self.panel.actions[0]['x'].set("10")
        self.panel._add_action_row() # idx 1
        self.panel.actions[1]['x'].set("30")
        
        # Insert after first (idx 0)
        self.panel._insert_action(0)
        self.assertEqual(len(self.panel.actions), 3)
        self.assertEqual(self.panel.actions[1]['x'].get(), "10") # Should inherit from row 0
        self.assertEqual(self.panel.actions[2]['x'].get(), "30")

if __name__ == '__main__':
    unittest.main()
