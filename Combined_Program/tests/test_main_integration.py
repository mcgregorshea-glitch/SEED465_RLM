import unittest
import sys
import os
import tkinter as tk
from unittest.mock import MagicMock, patch

_tests_dir = os.path.dirname(os.path.abspath(__file__))
_sc_root = os.path.dirname(_tests_dir)
_repo_root = os.path.dirname(_sc_root)
# src must come first so main.py resolves to seed-control-center/src/main.py, not legacy/
sys.path.insert(0, os.path.join(_sc_root, 'src'))
sys.path.insert(1, os.path.join(_repo_root, 'legacy', 'gui', 'src'))

from main import SEEDApplication

class TestMainIntegration(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        # Mock panels to avoid hardware connection attempts
        # We need to patch where they are IMPORTED in main.py
        with patch('main.PatternGeneratorGUI'), \
             patch('main.GCodeSenderGUI'), \
             patch('main.VivigoPanel') as MockVivigo:
            self.app = SEEDApplication(self.root)
            self.mock_vivigo = self.app.vivigo_panel

    def tearDown(self):
        self.root.destroy()

    def test_vivigo_tab_exists(self):
        """Verify the Vivigo Hub tab is added to the notebook (4 top-level tabs)."""
        tabs = self.app.notebook.tabs()
        # Exactly 4 top-level tabs: Generator, Manual, Sender, Vivigo Hub
        # POP Analysis is a subtab inside VivigoPanel, not a top-level tab
        self.assertEqual(len(tabs), 4)
        tab_texts = [self.app.notebook.tab(i, "text").strip() for i in range(len(tabs))]
        self.assertTrue(any("Vivigo" in text for text in tab_texts))

    def test_estop_disconnects_hub(self):
        """Verify E-STOP triggers Hub disconnection."""
        # Trigger the estop logic directly
        self.app.trigger_estop()
        
        self.mock_vivigo.disconnect_hub.assert_called_once()

if __name__ == '__main__':
    unittest.main()
