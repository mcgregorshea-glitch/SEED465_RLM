import unittest
from unittest.mock import MagicMock
import sys
import os

# Add src directory to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))
from sender_components.telemetry_service import DMMTelemetryProvider, VivigoTelemetryProvider

class TestTelemetryService(unittest.TestCase):
    def test_dmm_provider(self):
        mock_gui = MagicMock()
        mock_gui.dmm_group.read.return_value = [1.23]
        provider = DMMTelemetryProvider(mock_gui)
        data = provider.fetch_data()
        self.assertEqual(data, [1.23])

    def test_vivigo_provider(self):
        mock_gui = MagicMock()
        mock_gui.active_hub_modes = ["V_rec"]
        
        provider = VivigoTelemetryProvider(mock_gui)
        # Directly simulate a packet arrival
        provider.last_packet = {"v_rec": 12.0}
        
        data = provider.fetch_data()
        self.assertEqual(data, [12.0])

if __name__ == '__main__':
    unittest.main()
