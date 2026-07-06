"""
Serial read/write loop for the fake hub.
Runs on a daemon thread; communicates back to the UI via queues.
"""
import threading
import queue
import time

import serial

from hub_logic import HubState, TX_FREQ, RX_PACKET_SIZE
from hub_protocol import CobsDecoder


class HubServer:
    def __init__(self, port: str, baudrate: int = 115200):
        self.port      = port
        self.baudrate  = baudrate
        self.running   = False
        self._thread   = None

        self.state     = HubState()
        self.log_queue = queue.Queue()   # (str,) messages for the UI log
        self.connected = False

    # ── lifecycle ─────────────────────────────────────────────────────────────

    def start(self):
        self.running = True
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self):
        self.running = False
        if self._thread:
            self._thread.join(timeout=2.0)

    # ── main loop ─────────────────────────────────────────────────────────────

    def _run(self):
        while self.running:
            self._log(f"Opening {self.port} at {self.baudrate} baud...")
            try:
                with serial.Serial(self.port, self.baudrate, timeout=0) as port:
                    self.connected = True
                    self._log(f"Connected on {self.port}")
                    self._comm_loop(port)
            except serial.SerialException as e:
                self._log(f"Serial error: {e}")
            except Exception as e:
                self._log(f"Unexpected error in comms thread: {e}")
            finally:
                self.connected = False
                self._log("Disconnected.")
            if self.running:
                self._log("Retrying in 2 s...")
                time.sleep(2.0)

    def _comm_loop(self, port: serial.Serial):
        decoder    = CobsDecoder(RX_PACKET_SIZE)
        tx_interval = 1.0 / TX_FREQ
        next_tx    = time.monotonic()

        while self.running:
            now = time.monotonic()

            # ── RX: drain all waiting bytes ───────────────────────────────────
            n = port.in_waiting
            if n:
                for b in port.read(n):
                    packet = decoder.feed(b)
                    if packet is not None:
                        if self.state.apply_rx_packet(packet):
                            for msg in self.state.pop_log():
                                self._log(f"CMD: {msg}  [state -> {self._state_name()}]")
                        else:
                            self._log("RX: invalid packet (bad CRC or version)")

            # ── TX: send telemetry at TX_FREQ Hz ─────────────────────────────
            if now >= next_tx:
                next_tx = now + tx_interval
                try:
                    tx_bytes = self.state.build_tx_bytes()
                    port.write(tx_bytes)
                except Exception as e:
                    self._log(f"TX error (skipping frame): {e}")

            time.sleep(0.002)

    # ── helpers ───────────────────────────────────────────────────────────────

    def _log(self, msg: str):
        self.log_queue.put(msg)

    def _state_name(self) -> str:
        from hub_logic import STATE_NAMES
        return STATE_NAMES.get(self.state.state, '?')
