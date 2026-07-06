"""Regression tests for per-layer XY homing drift recovery.

Covers three defects:
  1. The ~30s stall: two batched 'ok's were consumed by the first of two
     chained _wait_for_ok() calls, starving the second until timeout.
  2. The endless verify/home loop: rewinding to the layer-start move without
     advancing past re-verification looped forever on persistent drift.
  3. "Resume at reduced speed" was computed and logged but never applied to
     the outgoing feedrate.

These exercise SerialEngine directly with a fake serial connection — no
tkinter, no hardware.
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


class _OkConn:
    """Marlin-ish fake: each '\\n' written queues one 'ok'. All pending oks are
    returned together on the next read (batched), like a real serial buffer when
    commands are written back-to-back."""
    def __init__(self):
        self._out = b""
        self.written = []
        self.lock_held_writes = []

    def write(self, data):
        self.written.append(data)
        self._out += b"ok\n" * data.count(b'\n')

    @property
    def in_waiting(self):
        return len(self._out)

    def read(self, n):
        d, self._out = self._out[:n], self._out[n:]
        return d


class _Event:
    def __init__(self, val=True):
        self._v = val
    def is_set(self):
        return self._v
    def wait(self, *a):
        return True


def _parse_coords(line):
    out = {}
    for ax, key in (('X', 'x'), ('Y', 'y'), ('Z', 'z'), ('E', 'rot')):
        m = re.search(rf'{ax}(-?\d+(?:\.\d+)?)', line.upper())
        if m:
            out[key] = float(m.group(1))
    return out


def _make_engine(conn):
    eng = SerialEngine.__new__(SerialEngine)
    eng.connection = conn
    eng.lock = threading.Lock()
    eng.message_queue = queue.Queue()
    eng.stop_event = _Event(False)
    eng.pause_event = _Event(True)
    eng.queue_message = lambda *a, **k: None
    # default callbacks
    eng.callbacks = {
        'take_measurement': None,
        'homing_verification': None,
        'handle_homing_failure': None,
        'get_hub_panel': None,
        'apply_e_conversion': lambda x: x,
        'parse_coords': _parse_coords,
        'get_speed_cap': lambda: getattr(eng, '_cap', 99999),
        'set_speed_cap': lambda x: setattr(eng, '_cap', x),
    }
    eng._cap = 99999
    return eng


class TestWaitForNOks(unittest.TestCase):
    def test_counts_batched_oks_without_starving(self):
        """Fix 1: the G28+M400 pair emits two batched oks; a single waiter must
        count both and return promptly instead of timing out."""
        eng = _make_engine(_OkConn())
        with eng.lock:
            eng.connection.write(b'G28 X Y\nM400\n')  # -> 'ok\nok\n'
        import time
        t0 = time.time()
        ok = eng._wait_for_n_oks(2, timeout_s=5)
        elapsed = time.time() - t0
        self.assertTrue(ok)
        self.assertLess(elapsed, 1.0, "should not stall waiting for a starved second ok")

    def test_returns_false_when_not_enough_oks(self):
        eng = _make_engine(_OkConn())
        with eng.lock:
            eng.connection.write(b'G28 X Y\n')  # only one ok
        self.assertFalse(eng._wait_for_n_oks(2, timeout_s=1))


class TestApplySpeedCap(unittest.TestCase):
    def setUp(self):
        self.eng = _make_engine(_OkConn())

    def test_no_cap_at_default_sentinel(self):
        self.eng._cap = 99999
        self.assertEqual(self.eng._apply_speed_cap("G1 X10 Y10 F3000"),
                         "G1 X10 Y10 F3000")

    def test_clamps_feedrate_above_cap(self):
        self.eng._cap = 1000
        self.assertEqual(self.eng._apply_speed_cap("G1 X10 Y10 F3000"),
                         "G1 X10 Y10 F1000")

    def test_leaves_feedrate_below_cap(self):
        self.eng._cap = 1000
        self.assertEqual(self.eng._apply_speed_cap("G1 X10 Y10 F500"),
                         "G1 X10 Y10 F500")

    def test_appends_cap_when_no_feedrate(self):
        self.eng._cap = 1000
        self.assertEqual(self.eng._apply_speed_cap("G1 X10 Y10"),
                         "G1 X10 Y10 F1000")


class TestDriftLoopRecovery(unittest.TestCase):
    """Fix 2 + 3, driven through the real _sender_thread."""

    def _bounds(self):
        return {'x_min': 0, 'x_max': 220, 'y_min': 0, 'y_max': 220,
                'z_min': 0, 'z_max': 128}

    def _layered_steps(self, n_layers):
        # First move sits at z_min (no verify on first layer change), then one
        # move per ascending layer. Layer change is detected when z changes and
        # the prior z is not z_min.
        steps = ["G1 X10 Y10 Z0 F3000"]
        for i in range(1, n_layers + 1):
            steps.append(f"G1 X{10+i} Y{10+i} Z{5*i} F3000")
        return steps

    def test_transient_drift_resumes_and_completes(self):
        """A single drift event re-homes, then the rewound boundary move resumes
        (no re-verify) and the scan runs to completion without hanging."""
        conn = _OkConn()
        eng = _make_engine(conn)
        calls = {'verify': 0, 'fail': 0}

        def verify(auto_restart=False):
            calls['verify'] += 1
            if calls['verify'] == 1:
                raise ValueError("Negative drift detected")  # one transient skip
            # subsequent verifications pass

        eng.callbacks['homing_verification'] = verify
        eng.callbacks['handle_homing_failure'] = lambda d: calls.__setitem__('fail', calls['fail'] + 1)

        steps = self._layered_steps(3)
        # Must terminate; guard against the historical infinite loop.
        done = threading.Event()

        def run():
            eng._sender_thread(steps, {'x': 0, 'y': 0, 'z': 0, 'rot': 0}, self._bounds(), False)
            done.set()

        t = threading.Thread(target=run, daemon=True)
        t.start()
        self.assertTrue(done.wait(timeout=10), "sender thread hung — drift loop did not terminate")
        self.assertEqual(calls['fail'], 0, "transient drift should not escalate to failure handler")
        # The blind recovery G28 X Y must have been issued exactly once.
        g28s = [w for w in conn.written if b'G28 X Y' in w]
        self.assertEqual(len(g28s), 1)
        # just_rehomed: the rewound boundary move must NOT be re-verified. Old
        # code re-verified it, giving 3 verify calls across the 3 layers; the
        # fix skips that one, leaving 2.
        self.assertEqual(calls['verify'], 2,
                         "rewound boundary move was re-verified (the loop bug)")
        # And the scan reached the final layer's move.
        sent = b"".join(conn.written)
        self.assertIn(b"Z15", sent)

    def test_persistent_drift_escalates_instead_of_looping(self):
        """If every verification fails, the cap halves to the floor and the run
        escalates to the failure handler rather than looping forever."""
        conn = _OkConn()
        eng = _make_engine(conn)
        calls = {'fail': 0}

        def verify(auto_restart=False):
            raise ValueError("Persistent drift")

        eng.callbacks['homing_verification'] = verify
        eng.callbacks['handle_homing_failure'] = lambda d: calls.__setitem__('fail', calls['fail'] + 1)

        steps = self._layered_steps(40)  # plenty of layers to drift on
        done = threading.Event()
        result = {}

        def run():
            try:
                eng._sender_thread(steps, {'x': 0, 'y': 0, 'z': 0, 'rot': 0}, self._bounds(), False)
            except InterruptedError as e:
                result['interrupted'] = str(e)
            done.set()

        t = threading.Thread(target=run, daemon=True)
        t.start()
        self.assertTrue(done.wait(timeout=10), "sender thread hung on persistent drift")
        self.assertGreaterEqual(calls['fail'], 1, "persistent drift must escalate to failure handler")
        # Cap was actually driven down to the floor region.
        self.assertLessEqual(eng._cap, 100)
        # Fix 3: the reduced cap must reach the wire — at least one move sent
        # with a clamped feedrate, not the original F3000.
        sent = b"".join(conn.written)
        self.assertTrue(re.search(rb'F(1000|500|250|125)\b', sent),
                        "speed cap was lowered but never applied to outgoing moves")


class TestPerLayerVerificationCoverage(unittest.TestCase):
    """Every completed layer must be drift-verified, except the very first
    (printer was just homed at scan start)."""

    def _bounds(self):
        return {'x_min': 0, 'x_max': 220, 'y_min': 0, 'y_max': 220,
                'z_min': 0, 'z_max': 128}

    def _run(self, z_sequence):
        conn = _OkConn()
        eng = _make_engine(conn)
        verified = []
        eng.callbacks['homing_verification'] = (
            lambda auto_restart=False: verified.append(True))
        steps = []
        for z in z_sequence:
            steps.append(f"G1 X10 Y10 Z{z} F3000")
            steps.append(f"G1 X20 Y10 Z{z} F3000")
        eng._sender_thread(steps, {'x': 0, 'y': 0, 'z': 0, 'rot': 0},
                           self._bounds(), False)
        return len(verified)

    def test_multi_rotation_revisits_zmin_still_verifies(self):
        """Regression: two rotations x two z-levels revisit the bottom (z_min)
        layer each rotation. The old `z == z_min` proxy skipped verification on
        every departure from the bottom layer, so only 1 of 3 boundaries was
        checked. Correct behaviour skips only the first boundary."""
        # z-sequence 0,5,0,5 -> 3 layer boundaries; skip the first, verify 2.
        self.assertEqual(self._run([0, 5, 0, 5]), 2)

    def test_single_rotation_offset_layers_all_verified(self):
        """Layers that never sit at z_min (typical center-offset scan) must all
        be verified except the first boundary."""
        # z-sequence 5,10,15,20 -> 3 boundaries; skip the first, verify 2.
        self.assertEqual(self._run([5, 10, 15, 20]), 2)


if __name__ == '__main__':
    unittest.main()
