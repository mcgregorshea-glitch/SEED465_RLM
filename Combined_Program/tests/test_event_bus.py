import unittest
import sys
import os

# Add src directory to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from utils import EventBus

class TestEventBus(unittest.TestCase):
    def test_event_bus_subscribe_publish(self):
        bus = EventBus()
        received_data = []

        def handler(data):
            received_data.append(data)

        bus.subscribe("test_event", handler)
        bus.publish("test_event", "test_data")

        self.assertEqual(received_data, ["test_data"])

if __name__ == '__main__':
    unittest.main()
