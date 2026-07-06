import unittest
import os
import sys
from unittest.mock import MagicMock, patch

_tests_dir = os.path.dirname(os.path.abspath(__file__))
_project_root = os.path.dirname(_tests_dir)
sys.path.insert(0, os.path.join(_project_root, 'src'))
sys.path.insert(0, os.path.join(_project_root, '..', 'legacy', 'gui', 'src'))

from sender_panel import GCodeSenderGUI

import tkinter as tk


class TestLoadToSender(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.parent = tk.Frame(self.root)
        self.gui = GCodeSenderGUI(self.parent)

    def tearDown(self):
        self.root.destroy()

    def _log_levels(self):
        levels = []
        for c in self.gui.log_message.call_args_list:
            if len(c.args) > 1:
                levels.append(c.args[1])
            else:
                levels.append(c.kwargs.get('level', 'INFO'))
        return levels

    @patch('sender_panel.GCodeSenderGUI._update_section_borders')
    def test_successful_load_populates_setup_pane(self, mock_borders):
        """Regression: 'Load to Sender' reaches the sender via load_gcode_file()
        directly, bypassing select_file(). The SETUP pane reads file_path_var /
        header_file_var, so a successful load must set them here."""
        path = os.path.join(os.sep, 'tmp', 'last_pattern.gcode')
        self.gui.file_path_var = MagicMock()
        self.gui.header_file_var = MagicMock()
        self.gui.log_message = MagicMock()

        def fake_process():
            self.gui.processed_gcode = [{'type': 'move'}, {'type': 'move'}]
            return True
        self.gui.process_gcode = fake_process

        self.gui.load_gcode_file(path)

        self.gui.file_path_var.set.assert_called_once_with(path)
        self.gui.header_file_var.set.assert_called_once_with(
            os.path.basename(path).upper())
        self.assertIn('SUCCESS', self._log_levels())

    @patch('sender_panel.messagebox.showerror')
    @patch('sender_panel.GCodeSenderGUI._update_section_borders')
    def test_failed_load_reports_error_not_success(self, mock_borders, mock_showerror):
        """A load that yields zero usable moves (e.g. out of bounds) must NOT be
        logged as SUCCESS; it must surface an error and not mark the file loaded."""
        path = os.path.join(os.sep, 'tmp', 'out_of_bounds.gcode')
        self.gui.file_path_var = MagicMock()
        self.gui.header_file_var = MagicMock()
        self.gui.log_message = MagicMock()

        def fake_process():
            self.gui.processed_gcode = []
            return False
        self.gui.process_gcode = fake_process

        self.gui.load_gcode_file(path)

        levels = self._log_levels()
        self.assertNotIn('SUCCESS', levels)
        self.assertIn('ERROR', levels)
        mock_showerror.assert_called_once()
        # SETUP pane must not advertise a loaded file.
        self.gui.file_path_var.set.assert_called_once_with('No file selected')


if __name__ == '__main__':
    unittest.main()
