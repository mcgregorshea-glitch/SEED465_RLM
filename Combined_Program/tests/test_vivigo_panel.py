import unittest
import sys
import os
import tkinter as tk
from unittest.mock import MagicMock, patch

_tests_dir = os.path.dirname(os.path.abspath(__file__))
_project_root = os.path.dirname(_tests_dir)
sys.path.insert(0, os.path.join(_project_root, 'src'))
sys.path.insert(0, os.path.join(_project_root, '..', 'legacy', 'gui', 'src'))

from vivigo_panel import VivigoPanel

class TestVivigoPanel(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.parent = tk.Frame(self.root)
        
        # Patch GUI class to avoid full initialization during testing
        with patch('vivigo_panel.GUI') as mock_gui:
            self.panel = VivigoPanel(self.parent)
            self.mock_gui_instance = mock_gui.return_value

    def tearDown(self):
        self.root.destroy()

    def test_panel_initialization(self):
        """Verify the panel initializes and creates the GUI frame."""
        self.assertIsInstance(self.panel, VivigoPanel)
        self.assertIsNotNone(self.panel.gui)
        self.assertEqual(self.panel.rlm, self.panel.gui)

    def test_disconnect_hub(self):
        """Verify the panel correctly triggers backend shutdown."""
        mock_backend = MagicMock()
        self.panel.gui.backend = mock_backend
        
        self.panel.disconnect_hub()
        mock_backend.end.assert_called_once()

    def test_connect_hub_compatibility(self):
        """Verify connect_hub calls port_selected on the embedded GUI."""
        self.panel.gui.port_selected = MagicMock()
        
        result = self.panel.connect_hub(None, "COM42")
        self.assertTrue(result)
        self.panel.gui.port_selected.assert_called_once_with("COM42")

    def test_set_param_delegation(self):
        """Verify set_param is delegated to the embedded GUI."""
        self.panel.gui.set_param = MagicMock(return_value=True)
        
        result = self.panel.rlm.set_param("test_param", 123)
        self.assertTrue(result)
        self.panel.gui.set_param.assert_called_once_with("test_param", 123)

if __name__ == '__main__':
    unittest.main()
