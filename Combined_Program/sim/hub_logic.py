"""
Hub state machine and telemetry generator.
Translates app.c + rlm_engi_params.h logic into Python.
Sinusoidal variation is added around the C firmware's fixed baseline values.
"""
import math
import time
import sys
import os
from dataclasses import dataclass, field

_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(_root, 'legacy', 'gui', 'src'))
sys.path.insert(0, os.path.join(_root, 'legacy'))
from backend import gen_fmt
from rlm_proj_vivigo import gen_project

from hub_protocol import (
    to_adc_vsys, to_adc_isys, to_adc_vbst, to_adc_ibst,
    to_adc_vrec, to_adc_tmp119,
    crc32_append, cobs_encode, CobsDecoder, crc32_verify,
)

# Hub firmware version (from app.c: RLM_FULL_VERSION)
HUB_VERSION    = 0x81207357
VERSION_MASK   = 0xFFFF0000   # upper 16 bits must match

# States (from rlm_states.h)
STATE_RESET    = 0
STATE_READY    = 1
STATE_SEARCH   = 2
STATE_LOCKED   = 3
STATE_SHUTDOWN = 4
STATE_FAULTED  = 5

STATE_NAMES = {
    STATE_RESET:    'Reset',
    STATE_READY:    'WPT Ready',
    STATE_SEARCH:   'WPT Search',
    STATE_LOCKED:   'WPT Locked',
    STATE_SHUTDOWN: 'Shutdown',
    STATE_FAULTED:  'Faulted',
}

STATE_COLORS = {
    STATE_RESET:    '#888888',
    STATE_READY:    '#4CAF50',
    STATE_SEARCH:   '#FF9800',
    STATE_LOCKED:   '#2196F3',
    STATE_SHUTDOWN: '#9C27B0',
    STATE_FAULTED:  '#F44336',
}

# TX / RX packet sizes (from rlm_engi_params.h)
TX_PACKET_SIZE = 76   # 72 bytes data + 4 bytes CRC-32
RX_PACKET_SIZE = 9    # 5 bytes data  + 4 bytes CRC-32

TX_FREQ = 5           # Hz (from RLM_ENGI_TX_FREQ in app.c)

# Baseline telemetry values (from app.c:tx_tick fixed assignments)
_BASE = dict(
    v_sys=11.1, i_sys=0.400,
    v_bst_idle=40.0, v_bst_locked=80.0,
    i_bst_idle=0.200, i_bst_locked=0.050,
    t_bst=30.0, t_inv=35.0,
    v_rec=4.5, t_rec=25.0,
)

# Sinusoidal variation defaults: (amplitude, period_s)
# Thermal channels use slow periods to feel realistic.
_SIN = dict(
    v_sys=(0.4,  5.0),
    i_sys=(0.04, 3.1),
    v_bst=(2.0,  2.0),
    i_bst=(0.020, 2.3),
    t_bst=(1.5, 12.0),
    t_inv=(1.5, 11.3),
    v_rec=(0.15,  4.0),
    t_rec=(0.8,  14.0),
)

# When LOCKED, amplitudes scale up to simulate active WPT power delivery.
_LOCKED_AMP_SCALE = 2.5


@dataclass
class SineParams:
    amplitude: float
    period: float    # seconds (= 1 / frequency)
    phase: float = 0.0  # radians


# Mutable runtime params — starts as a copy of defaults, modified by the GUI.
_SINE_PARAMS: dict[str, SineParams] = {
    key: SineParams(amp, period)
    for key, (amp, period) in _SIN.items()
}


def get_sine_params(key: str) -> SineParams:
    return _SINE_PARAMS[key]


def set_sine_params(key: str, params: SineParams) -> None:
    _SINE_PARAMS[key] = params


def get_all_sine_params() -> dict[str, SineParams]:
    return dict(_SINE_PARAMS)


def load_sine_params(data: dict) -> None:
    """Load params from a dict of {key: {amplitude, period, phase}}."""
    for key, vals in data.items():
        if key in _SINE_PARAMS:
            _SINE_PARAMS[key] = SineParams(
                amplitude=float(vals['amplitude']),
                period=float(vals['period']),
                phase=float(vals.get('phase', 0.0)),
            )


def reset_sine_params() -> None:
    """Restore all params to _SIN hardcoded defaults."""
    for key, (amp, period) in _SIN.items():
        _SINE_PARAMS[key] = SineParams(amp, period)


def _sin(key: str, t: float, locked: bool) -> float:
    p = _SINE_PARAMS[key]
    amp = p.amplitude * (_LOCKED_AMP_SCALE if locked else 1.0)
    return amp * math.sin(2 * math.pi * t / p.period + p.phase)


def _build_fmt():
    project = gen_project()
    board   = project.boards[1]
    return gen_fmt(board.values_board()), gen_fmt(board.params_board())

_tx_fmt, _rx_fmt = _build_fmt()


class HubState:
    def __init__(self):
        self.state = STATE_READY
        self.button_pressed = False
        self._start = time.monotonic()
        self._last_cmd_log: list[str] = []   # consumed by UI each frame

    # ── command handling ──────────────────────────────────────────────────────

    def apply_rx_packet(self, packet: bytes) -> bool:
        """
        Decode a 9-byte RX packet (after COBS, before CRC strip).
        Returns True if the packet was valid and processed.
        """
        if not crc32_verify(packet):
            return False
        data = _rx_fmt.unpack(packet[:5])
        version = data['Hub Version']
        if (version & VERSION_MASK) != (HUB_VERSION & VERSION_MASK):
            return False

        wpt_start = bool(data['WPT Start'])
        wpt_stop  = bool(data['WPT Stop'])

        if wpt_start:
            self.state = STATE_LOCKED
            self._last_cmd_log.append('WPT_START received')
        if wpt_stop:
            self.state = STATE_READY
            self._last_cmd_log.append('WPT_STOP received')
        return True

    def pop_log(self) -> list[str]:
        msgs, self._last_cmd_log = self._last_cmd_log, []
        return msgs

    # ── telemetry ─────────────────────────────────────────────────────────────

    def build_tx_bytes(self) -> bytes:
        """Build the 76-byte COBS-framed telemetry packet ready to write to serial."""
        payload = self._pack_payload()
        with_crc = crc32_append(payload)
        return cobs_encode(with_crc)

    def _pack_payload(self) -> bytes:
        t = time.monotonic() - self._start
        locked = (self.state == STATE_LOCKED)

        v_sys  = _BASE['v_sys']  + _sin('v_sys',  t, locked)
        i_sys  = _BASE['i_sys']  + _sin('i_sys',  t, locked)
        t_bst  = _BASE['t_bst']  + _sin('t_bst',  t, locked)
        t_inv  = _BASE['t_inv']  + _sin('t_inv',  t, locked)
        v_rec  = _BASE['v_rec']  + _sin('v_rec',  t, locked)
        t_rec  = _BASE['t_rec']  + _sin('t_rec',  t, locked)

        if self.button_pressed or locked:
            v_bst = _BASE['v_bst_locked'] + _sin('v_bst', t, locked)
            i_bst = _BASE['i_bst_locked'] + _sin('i_bst', t, locked)
        else:
            v_bst = _BASE['v_bst_idle'] + _sin('v_bst', t, locked)
            i_bst = _BASE['i_bst_idle'] + _sin('i_bst', t, locked)

        v_sys_adc  = max(0, min(0xFFF, to_adc_vsys(v_sys)))
        i_sys_adc  = max(0, min(0xFFF, to_adc_isys(i_sys)))
        v_bst_adc  = max(0, min(0xFFF, to_adc_vbst(v_bst)))
        i_bst_adc  = max(0, min(0xFFF, to_adc_ibst(i_bst)))
        t_bst_adc  = max(0, min(0xFFF, to_adc_tmp119(t_bst)))
        t_inv_adc  = max(0, min(0xFFF, to_adc_tmp119(t_inv)))
        v_rec_adc  = max(0, min(0xFFF, to_adc_vrec(v_rec)))
        t_rec_adc  = max(0, min(0xFFF, to_adc_tmp119(t_rec)))

        fields = {
            # ── ADC samples ──────────────────────────────────────────────────
            'v_sys_smp': v_sys_adc, 'i_sys_smp': i_sys_adc,
            'v_bst_smp': v_bst_adc, 'i_bst_smp': i_bst_adc,
            't_bst_smp': t_bst_adc, 't_inv_smp': t_inv_adc,
            'v_rec_smp': v_rec_adc, 't_rec_smp': t_rec_adc,
            # ── avg valid flags (all valid) ──────────────────────────────────
            'v_sys_avg_valid': 1, 'i_sys_avg_valid': 1,
            'v_bst_avg_valid': 1, 'i_bst_avg_valid': 1,
            't_bst_avg_valid': 1, 't_inv_avg_valid': 1,
            'v_rec_avg_valid': 1, 't_rec_avg_valid': 1,
            # ── averages (same as samples; no hardware averaging in sim) ─────
            'v_sys_avg': v_sys_adc, 'i_sys_avg': i_sys_adc,
            'v_bst_avg': v_bst_adc, 'i_bst_avg': i_bst_adc,
            't_bst_avg': t_bst_adc, 't_inv_avg': t_inv_adc,
            'v_rec_avg': v_rec_adc, 't_rec_avg': t_rec_adc,
            # ── trip latch (all clear) ───────────────────────────────────────
            'v_sys_tripped': 0, 'i_sys_tripped': 0,
            'v_bst_tripped': 0, 'i_bst_tripped': 0,
            't_bst_tripped': 0, 't_inv_tripped': 0,
            'v_rec_tripped': 0, 't_rec_tripped': 0,
            # ── board values ─────────────────────────────────────────────────
            'State':              self.state,
            'Trips and Commands': 0,
            'Fault Count':        0,
            'WPT Enabled':        1 if locked else 0,
            'FPGA Buck Deadtime': 0,
            'FPGA Boost Deadtime': 0,
            'FPGA A Coarse':      0,
            'FPGA A Fine':        0,
            'FPGA B Coarse':      0,
            'FPGA B Fine':        0,
            'Hub SOC':            0xC00,
            'Implant SOC':        0x7FF,
            'UI Button':          1 if self.button_pressed else 0,
            'UI Counter':         0,
            'USB Charging':       0,
            'Alignment':          0xA0,
            'PETV CTRL':          0,
            'Control Error':      0,
            'Control Voltage':    0,
            'Implant Synced':     1,
            'Implant CMD Stop':   0,
            'Implant State':      3,
            'Implant Version':    0x01237357,
            'BMS PG':             1,
            'BMS INT':            0,
            'BMS STAT0':          0, 'BMS STAT1': 0, 'BMS STAT2': 0,
            'BMS FLAG0':          0, 'BMS FLAG1': 0, 'BMS FLAG2': 0, 'BMS FLAG3': 0,
            'PETV CTRL Valids':   200,
            'PETV Slow Valids':   13,
            'PETV Stats Value':   0,
            'PETV Resp Valids':   0,
            'PETV Resp Valids IDX': 0,
            # ── echo-back controllable params (init values from rlm_engi_params.h) ──
            'Open Loop':                0,
            'DCDC Voltage':             26,
            'Control kI':               123,
            'PETV Compare Threshold':   33,
            'PETV Stats Index':         1,
            # ── params_board echo ────────────────────────────────────────────
            'WPT Start':    0,
            'WPT Stop':     0,
            'Hub Version':  HUB_VERSION,
        }
        return bytes(_tx_fmt.pack(fields))

    # ── real-world values for the UI readout ──────────────────────────────────

    def get_real_values(self) -> dict:
        """Returns the current real-world telemetry values (for GUI display)."""
        t = time.monotonic() - self._start
        locked = (self.state == STATE_LOCKED)

        v_sys = _BASE['v_sys'] + _sin('v_sys', t, locked)
        i_sys = _BASE['i_sys'] + _sin('i_sys', t, locked)
        t_bst = _BASE['t_bst'] + _sin('t_bst', t, locked)
        t_inv = _BASE['t_inv'] + _sin('t_inv', t, locked)
        v_rec = _BASE['v_rec'] + _sin('v_rec', t, locked)
        t_rec = _BASE['t_rec'] + _sin('t_rec', t, locked)
        if self.button_pressed or locked:
            v_bst = _BASE['v_bst_locked'] + _sin('v_bst', t, locked)
            i_bst = _BASE['i_bst_locked'] + _sin('i_bst', t, locked)
        else:
            v_bst = _BASE['v_bst_idle'] + _sin('v_bst', t, locked)
            i_bst = _BASE['i_bst_idle'] + _sin('i_bst', t, locked)

        return dict(v_sys=v_sys, i_sys=i_sys, v_bst=v_bst, i_bst=i_bst,
                    t_bst=t_bst, t_inv=t_inv, v_rec=v_rec, t_rec=t_rec)
