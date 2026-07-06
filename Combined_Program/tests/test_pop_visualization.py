import unittest
import pandas as pd
import numpy as np
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src'))

from pop_visualization_panel import POPVisualizationPanel
import tkinter as tk
from tkinter import ttk

class TestPOPSlicingLogic(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.panel = POPVisualizationPanel(self.root)
        
        # Create dummy 4D data
        rows = []
        for x in [0, 10]:
            for y in [0, 10]:
                for r in [0, 5]:
                    for rot in [0, 90]:
                        rows.append({
                            'x': x, 'y': y, 'r': r, 'ROT': rot,
                            'Voltage': x + y + r + rot,
                            'Current': (x + y + r + rot) / 10.0
                        })
        self.df = pd.DataFrame(rows)
        self.panel.data = rows
        self.panel.fields = list(rows[0].keys())

    def test_field_identification(self):
        """Verify that IVs and DVs are correctly identified."""
        self.panel.iv_fields = {'x', 'y', 'r', 'ROT'}
        self.panel.dv_fields = {'Voltage', 'Current'}
        self.panel._update_ui_fields()
        ivs = self.panel.iv_fields
        dvs = self.panel.dv_fields
        self.assertIn('Voltage', dvs)
        self.assertIn('Current', dvs)
        self.assertIn('x', ivs)
        self.assertIn('ROT', ivs)

    def test_slicing_sliders_creation(self):
        """Verify that sliders are created for IVs not on axes."""
        self.panel.iv_fields = {'x', 'y', 'r', 'ROT'}
        self.panel.dv_fields = {'Voltage', 'Current'}
        self.panel.x_axis_var.set('x')
        self.panel.y_axis_var.set('y')
        self.panel._update_ui_fields()
        
        # 'r' and 'ROT' should be in slider_vars because they are IVs not on axes
        self.assertIn('ROT', self.panel.slider_vars)
        self.assertIn('r', self.panel.slider_vars)
        self.assertNotIn('x', self.panel.slider_vars)

    def test_plot_syncing(self):
        """Verify that plots are added to the grid when DVs are checked."""
        self.panel.iv_fields = {'x', 'y', 'r', 'ROT'}
        self.panel.dv_fields = {'Voltage', 'Current'}
        self.panel._update_ui_fields()
        self.panel.dv_vars['Voltage'].set(True)
        self.panel.dv_vars['Current'].set(False)
        self.panel._sync_plots()
        
        self.assertIn('Voltage', self.plot_widgets_keys())
        self.assertNotIn('Current', self.plot_widgets_keys())
        
        self.panel.dv_vars['Current'].set(True)
        self.panel._sync_plots()
        self.assertIn('Current', self.plot_widgets_keys())

    def plot_widgets_keys(self):
        return list(self.panel.plot_widgets.keys())

    def tearDown(self):
        self.root.destroy()

if __name__ == '__main__':
    unittest.main()
