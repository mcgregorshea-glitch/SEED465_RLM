import unittest
import sys
import os
import time
from unittest.mock import MagicMock, patch, PropertyMock

_tests_dir = os.path.dirname(os.path.abspath(__file__))
_project_root = os.path.dirname(_tests_dir)
sys.path.insert(0, os.path.join(_project_root, 'src'))
sys.path.insert(0, os.path.join(_project_root, '..', 'legacy', 'gui', 'src'))

from sender_panel import GCodeSenderGUI

import tkinter as tk

class TestSequenceIntegration(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.parent = tk.Frame(self.root)
        
        # Mock serial and RLM
        self.mock_serial = MagicMock()
        self.mock_rlm = MagicMock()
        
        # Instantiate GCodeSenderGUI
        # We need to mock things that try to talk to hardware in __init__ if any
        with patch('sender_panel.serial.Serial', return_value=self.mock_serial):
            self.gui = GCodeSenderGUI(self.parent)
            self.gui.serial_connection = self.mock_serial
            # Initialize coordinates to avoid None issues in the sender thread
            self.gui.last_cmd_abs_x = 0.0
            self.gui.last_cmd_abs_y = 0.0
            self.gui.last_cmd_abs_z = 0.0
            self.gui.last_cmd_abs_rot = 0.0
            self.gui.is_connected = True
            # For testing, we'll manually set the rlm_instance
            self.gui.rlm_instance = self.mock_rlm

    def tearDown(self):
        self.root.destroy()

    @patch('sender_panel.messagebox.askyesno', return_value=True)
    @patch('sender_panel.messagebox.showerror')
    @patch('sender_panel.messagebox.showwarning')
    @patch('sender_panel.messagebox.showinfo')
    @patch('sender_panel.GCodeSenderGUI._show_scan_complete_popup')
    @patch('sender_panel.GCodeSenderGUI.process_gcode')
    @patch('sender_panel.GCodeSenderGUI._update_all_displays')
    def test_run_sequence_basic(self, mock_update, mock_process, mock_popup, mock_info, mock_warning, mock_error, mock_ask):
        """Verify that run_sequence executes motion and hub commands."""
        sequence = [
            {"type": "move", "gcode": "G1 X10 Y10"},
            {"type": "hub_cmd", "param": "WPT Start", "value": True},
            {"type": "wait", "seconds": 0.01}
        ]
        
        # Setup mocks
        self.gui.vivigo_panel = MagicMock()
        self.gui.vivigo_panel.rlm = self.mock_rlm
        
        # Manually set processed_gcode and log path for the test
        self.gui.processed_gcode = sequence
        self.gui.log_filepath_var.set("test_log.csv")
        self.gui.rotation_crash_test_complete = True
        
        # mock serial to return ok when checked
        self.mock_serial.in_waiting = 3
        self.mock_serial.read.return_value = b"ok\n"
        
        # Ensure the mock serial always reports data waiting
        type(self.mock_serial).in_waiting = PropertyMock(return_value=3)

        # Set the connection on the engine too
        self.gui.serial_engine.connection = self.mock_serial
        
        # Disable auto-measure to avoid infinite gated measurement loop in test
        self.gui.auto_measure_enabled.set(False)

        # Patch threading.Thread to run synchronously for testing
        with patch('threading.Thread') as mock_thread:
            # Capture the target function
            def create_sync_thread(*args, **kwargs):
                target = kwargs.get('target') or (args[0] if args else None)
                target_args = kwargs.get('args') or (args[1] if len(args) > 1 else [])
                
                mock_t = MagicMock()
                mock_t.start.side_effect = lambda: target(*target_args) if target else None
                return mock_t
            mock_thread.side_effect = create_sync_thread
            
            with patch('time.sleep'):
                self.gui.start_sending()

        # 1. Hub command
        self.gui.vivigo_panel.rlm.set_param.assert_called_with("WPT Start", True)

        # 2. Verify G-Code was sent
        sent_commands = [call.args[0].decode('utf-8').strip() for call in self.mock_serial.write.call_args_list]
        self.assertIn("G1 X10 Y10", sent_commands)

if __name__ == '__main__':
    unittest.main()
