import unittest
from unittest.mock import MagicMock, patch
import sys
import os
import time

# Add src directory to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from sender_panel import GCodeSenderGUI

class TestPauseCapture(unittest.TestCase):
    def setUp(self):
        # Create a mock for GCodeSenderGUI
        self.sender = MagicMock(spec=GCodeSenderGUI)
        # We'll use the actual method by attaching it to the mock
        from sender_panel import GCodeSenderGUI as ActualSender
        self.sender._perform_automatic_capture = ActualSender._perform_automatic_capture.__get__(self.sender)
        self.sender._on_measurement = ActualSender._on_measurement.__get__(self.sender)
        self.sender._get_center = ActualSender._get_center.__get__(self.sender)
        self.sender._to_relative = ActualSender._to_relative.__get__(self.sender)
        
        self.sender.pre_measure_delay_var = MagicMock()
        self.sender.pre_measure_delay_var.get.return_value = 0.2
        self.sender.data_source_var = MagicMock()
        self.sender.data_source_var.get.return_value = "Hub"
        self.sender.active_data_source = "Hub"
        
        self.sender.hub_mode_vars = {"V_rec": MagicMock()}
        self.sender.hub_modes_list = ["V_rec"]
        self.sender.active_hub_modes = ["V_rec"]
        
        # Center vars
        self.sender.center_x_var = MagicMock(); self.sender.center_x_var.get.return_value = "0.0"
        self.sender.center_y_var = MagicMock(); self.sender.center_y_var.get.return_value = "0.0"
        self.sender.center_z_var = MagicMock(); self.sender.center_z_var.get.return_value = "0.0"
        self.sender.center_rot_var = MagicMock(); self.sender.center_rot_var.get.return_value = "0.0"
        
        # Mock dependencies for _perform_automatic_capture
        self.sender.vivigo_telemetry = MagicMock()
        self.sender.vivigo_telemetry.fetch_data.return_value = [25.0]
        
        self.sender._log_measurement_to_file = MagicMock()
        self.sender.event_bus = MagicMock()
        self.sender.message_queue = MagicMock()
        
        self.sender.last_measurement_var = MagicMock()
        self.sender.log_measurements_enabled = MagicMock()
        self.sender.log_measurements_enabled.get.return_value = True
        
        self.sender.last_cmd_abs_x = 10.0
        self.sender.last_cmd_abs_y = 20.0
        self.sender.last_cmd_abs_z = 30.0
        self.sender.last_cmd_abs_rot = 0.0

    def test_on_measurement_processing(self):
        # Act - Simulate a measurement result coming from the SerialEngine queue
        vals = [25.0]
        coords = {'x': 10.0, 'y': 20.0, 'z': 30.0, 'rot': 5.0}
        self.sender._on_measurement(vals, coords)
        
        # Assert
        self.sender._log_measurement_to_file.assert_called_once()
        self.sender.event_bus.publish.assert_any_call("gated_measurement", {
            'x': 10.0, 'y': 20.0, 'z': 30.0, 'rot': 5.0, 'V_rec': 25.0
        })

if __name__ == '__main__':
    unittest.main()
