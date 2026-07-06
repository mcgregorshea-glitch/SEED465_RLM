import unittest
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src'))

from sequence_manager import SequenceManager

class TestSequenceManager(unittest.TestCase):
    def setUp(self):
        self.manager = SequenceManager()
        self.test_file = "test_sequence.json"

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def test_load_valid_sequence(self):
        """Verify that a valid JSON sequence is loaded correctly."""
        sequence_data = [
            {"type": "move", "x": 10, "y": 20},
            {"type": "hub_cmd", "param": "WPT Start", "value": True},
            {"type": "wait", "seconds": 5}
        ]
        with open(self.test_file, 'w') as f:
            json.dump(sequence_data, f)

        loaded_sequence = self.manager.load_sequence(self.test_file)
        self.assertEqual(len(loaded_sequence), 3)
        self.assertEqual(loaded_sequence[0]['type'], 'move')
        self.assertEqual(loaded_sequence[1]['param'], 'WPT Start')

    def test_invalid_json(self):
        """Verify that malformed JSON raises a ValueError."""
        with open(self.test_file, 'w') as f:
            f.write("invalid json content")
        
        with self.assertRaises(ValueError):
            self.manager.load_sequence(self.test_file)

    def test_invalid_sequence_structure(self):
        """Verify that a non-list sequence raises a TypeError."""
        sequence_data = {"type": "not_a_list"}
        with open(self.test_file, 'w') as f:
            json.dump(sequence_data, f)
            
        with self.assertRaises(TypeError):
            self.manager.load_sequence(self.test_file)

if __name__ == '__main__':
    unittest.main()
