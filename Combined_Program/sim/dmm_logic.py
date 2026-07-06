"""
Simulated measurement state for a single SCPI multimeter.
Mirrors the CONF / SAMP:COUN / CALC:AVER:* sequence that dmm_manager.py
(DmmInst.setup/trigger/ready/read) drives against a real instrument.
"""
import random
import time


class DmmState:
    def __init__(self, idn: str = "Simulated Instruments,MODEL-34461A,SN00000001,FW1.00"):
        self.idn = idn
        self.mode = "VOLT:DC"
        self.sample_target = 10
        self.averaging_enabled = False
        self.aver_count = 0
        self._trigger_time: float | None = None

        # User-adjustable via the GUI: the "true" value the instrument reports,
        # plus +/- percent noise applied on every CALC:AVER:ALL? read.
        self.base_value = 5.000
        self.noise_pct = 0.5
        self.samples_per_sec = 20.0

    def configure(self, mode: str) -> None:
        self.mode = mode.strip()

    def set_sample_count(self, n: int) -> None:
        self.sample_target = max(1, n)

    def set_averaging(self, on: bool) -> None:
        self.averaging_enabled = on

    def trigger(self) -> None:
        self.aver_count = 0
        self._trigger_time = time.monotonic()

    def current_aver_count(self) -> int:
        if self._trigger_time is None:
            return self.aver_count
        elapsed = time.monotonic() - self._trigger_time
        self.aver_count = min(int(elapsed * self.samples_per_sec), self.sample_target)
        return self.aver_count

    def averaged_reading(self) -> float:
        noise = self.base_value * (self.noise_pct / 100.0) * (random.random() * 2 - 1)
        return self.base_value + noise
