"""Measurements must be logged against the printer's ACTUAL position (M114),
not the commanded target.

Regression for the X=210/200 data-integrity bug: the firmware clamped an
out-of-range X target (210 -> 200 travel limit) and the printer reported 200,
but the scan logged the commanded 210. The engine now reads M114 at each
measurement so the recorded coordinate matches physical reality on every axis
(a clamp or skipped steps can no longer be silently mislogged).
"""
import os
import re
import sys
import queue
import threading
import unittest

_tests_dir = os.path.dirname(os.path.abspath(__file__))
_project_root = os.path.dirname(_tests_dir)
sys.path.insert(0, os.path.join(_project_root, 'src'))

from sender_components.serial_engine import SerialEngine


class _ClampConn:
    """Marlin-ish fake that clamps X/Y targets to a max travel limit (the
    firmware bug) and reports the clamped position through M114, like the real
    printer's LCD readout."""
    def __init__(self, max_travel=200.0):
        self._out = b""
        self.max = max_travel
        self.pos = {'x': 0.0, 'y': 0.0, 'z': 0.0, 'e': 0.0}

    def write(self, data):
        line = data.decode('utf-8', errors='ignore').strip()
        if line.startswith('M114'):
            p = self.pos
            self._out += (f"X:{p['x']:.2f} Y:{p['y']:.2f} Z:{p['z']:.2f} "
                          f"E:{p['e']:.2f} Count X:0 Y:0 Z:0\nok\n").encode()
            return
        if line.startswith('G0') or line.startswith('G1'):
            for ax, key in (('X', 'x'), ('Y', 'y'), ('Z', 'z'), ('E', 'e')):
                m = re.search(rf"{ax}([-+]?\d*\.?\d+)", line)
                if m:
                    v = float(m.group(1))
                    if key in ('x', 'y'):
                        v = min(v, self.max)  # firmware travel clamp
                    self.pos[key] = v
        self._out += b"ok\n"

    @property
    def in_waiting(self):
        return len(self._out)

    def read(self, n):
        d, self._out = self._out[:n], self._out[n:]
        return d

    def flush(self):
        pass


def _parse_coords(line):
    out = {}
    for ax, key in (('X', 'x'), ('Y', 'y'), ('Z', 'z'), ('E', 'rot')):
        m = re.search(rf'{ax}(-?\d+(?:\.\d+)?)', line.upper())
        if m:
            out[key] = float(m.group(1))
    return out


class _Event:
    def __init__(self, val=True): self._v = val
    def is_set(self): return self._v
    def wait(self, *a): return True


def _make_engine(conn):
    eng = SerialEngine.__new__(SerialEngine)
    eng.connection = conn
    eng.lock = threading.Lock()
    eng.message_queue = queue.Queue()
    eng.stop_event = _Event(False)
    eng.pause_event = _Event(True)
    eng.queue_message = lambda *a, **k: None
    eng.callbacks = {
        'take_measurement': lambda: [1.23],
        'homing_verification': None,
        'handle_homing_failure': None,
        'get_hub_panel': None,
        'apply_e_conversion': lambda x: x,
        'parse_coords': _parse_coords,
        'get_speed_cap': lambda: 99999,
        'set_speed_cap': lambda x: None,
    }
    return eng


def _measured_positions(eng):
    out = []
    while not eng.message_queue.empty():
        kind, payload = eng.message_queue.get()
        if kind == 'MEASUREMENT_RESULT':
            out.append(payload[1])
    return out


class TestActualPositionLogging(unittest.TestCase):
    def _bounds(self):
        return {'x_min': 0, 'x_max': 200, 'y_min': 0, 'y_max': 200,
                'z_min': 0, 'z_max': 128}

    def test_clamped_axis_logs_actual_not_commanded(self):
        eng = _make_engine(_ClampConn(max_travel=200.0))
        steps = ["G1 X210.000 Y150.000 Z5.000 E0.000 F1000"]
        eng._sender_thread(steps, {'x': 0, 'y': 0, 'z': 0, 'rot': 0},
                           self._bounds(), True)
        pos = _measured_positions(eng)
        self.assertEqual(len(pos), 1)
        # Commanded X210 was clamped to 200 by firmware; log must record 200.
        self.assertAlmostEqual(pos[0]['x'], 200.0, places=2)
        # Y was within range and must be recorded unchanged.
        self.assertAlmostEqual(pos[0]['y'], 150.0, places=2)
        self.assertAlmostEqual(pos[0]['z'], 5.0, places=2)

    def test_in_range_move_logs_true_position(self):
        eng = _make_engine(_ClampConn(max_travel=200.0))
        steps = ["G1 X75.000 Y40.000 Z10.000 E0.000 F1000"]
        eng._sender_thread(steps, {'x': 0, 'y': 0, 'z': 0, 'rot': 0},
                           self._bounds(), True)
        pos = _measured_positions(eng)
        self.assertEqual(len(pos), 1)
        self.assertAlmostEqual(pos[0]['x'], 75.0, places=2)
        self.assertAlmostEqual(pos[0]['y'], 40.0, places=2)

    def test_falls_back_to_commanded_when_m114_unparseable(self):
        """If the M114 reply can't be parsed, logging must not break — it falls
        back to the commanded target rather than dropping the measurement."""
        class _NoPosConn(_ClampConn):
            def write(self, data):
                line = data.decode('utf-8', errors='ignore').strip()
                if line.startswith('M114'):
                    self._out += b"ok\n"   # no X:/Y:/Z: tokens
                    return
                super().write(data)
        eng = _make_engine(_NoPosConn(max_travel=200.0))
        steps = ["G1 X50.000 Y60.000 Z7.000 E0.000 F1000"]
        eng._sender_thread(steps, {'x': 0, 'y': 0, 'z': 0, 'rot': 0},
                           self._bounds(), True)
        pos = _measured_positions(eng)
        self.assertEqual(len(pos), 1)
        self.assertAlmostEqual(pos[0]['x'], 50.0, places=2)
        self.assertAlmostEqual(pos[0]['y'], 60.0, places=2)


if __name__ == '__main__':
    unittest.main()
