import sys
import os
import unittest
from unittest.mock import MagicMock, patch

# legacy/ lives at the repo root, two levels up from this test file
_tests_dir = os.path.dirname(os.path.abspath(__file__))
_sc_root = os.path.dirname(_tests_dir)        # seed-control-center/
_repo_root = os.path.dirname(_sc_root)        # repo root (RLM/)
gui_src_dir = os.path.join(_repo_root, "legacy", "gui", "src")
seed_gui_dir = os.path.join(_repo_root, "legacy")

# Use insert(0, ...) so these paths win over pytest's rootdir resolution
if gui_src_dir not in sys.path:
    sys.path.insert(0, gui_src_dir)
if seed_gui_dir not in sys.path:
    sys.path.insert(0, seed_gui_dir)

class TestRLMAPIInit(unittest.TestCase):
    def test_api_import_and_instantiation(self):
        """Verifies that the RLM API can be imported and instantiated with a config."""
        try:
            from rlm_api import RLM
            import rlm_proj_vivigo
            
            config_path = os.path.join(seed_gui_dir, "rlm_proj_vivigo.py")
            rlm = RLM(config_path)
            self.assertIsNotNone(rlm)
            print("PASS: RLM API initialized successfully.")
            
        except ImportError as e:
            self.fail(f"Failed to import RLM API dependencies: {e}")
        except Exception as e:
            self.fail(f"Unexpected error during RLM API initialization: {e}")

    @patch('backend.Backend')
    def test_api_mock_run(self, mock_backend):
        """Verifies that the API can 'run' with a mocked backend."""
        from rlm_api import RLM
        config_path = os.path.join(seed_gui_dir, "rlm_proj_vivigo.py")
        try:
            rlm = RLM(config_path)
            self.assertIsNotNone(rlm)
            # With Backend mocked, no real serial port is opened
            rlm.api_run("COM_MOCK", console_log=False)
            import time
            time.sleep(0.1)
            rlm.api_end()
            print("PASS: RLM API run loop (mocked) completed without crashing.")
        except Exception as e:
            self.fail(f"RLM API run loop failed: {e}")

if __name__ == "__main__":
    unittest.main()
