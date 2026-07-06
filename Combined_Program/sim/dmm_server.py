"""
TCP/SCPI transport for the fake DMM.

Real DMMs here are addressed over the network via PyVISA (see
sender_components/dmm_manager.py: DmmInst.connect() tries a VXI-11 resource
first, then falls back to `TCPIP::<ip>::5025::SOCKET` — a plain ASCII SCPI
socket). This server implements just that fallback: one line in, one
optional line out, newline-terminated.
"""
import queue
import socket
import threading

from dmm_logic import DmmState


class DmmServer:
    def __init__(self, host: str, port: int = 5025):
        self.host = host
        self.port = port
        self.running = False
        self._thread = None
        self._sock: socket.socket | None = None

        self.state = DmmState()
        self.log_queue: queue.Queue = queue.Queue()
        self.connected = False

    # ── lifecycle ────────────────────────────────────────────────────────────

    def start(self):
        self.running = True
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self):
        self.running = False
        if self._sock:
            try:
                self._sock.close()
            except OSError:
                pass
        if self._thread:
            self._thread.join(timeout=2.0)

    # ── main loop ────────────────────────────────────────────────────────────

    def _run(self):
        try:
            self._sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self._sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self._sock.bind((self.host, self.port))
            self._sock.listen(1)
            self._sock.settimeout(0.5)
            self._log(f"Listening on {self.host}:{self.port}")
        except OSError as e:
            self._log(f"Bind failed on {self.host}:{self.port}: {e}")
            self.running = False
            return

        while self.running:
            try:
                conn, addr = self._sock.accept()
            except socket.timeout:
                continue
            except OSError:
                break

            self.connected = True
            self._log(f"Client connected from {addr[0]}:{addr[1]}")
            try:
                self._comm_loop(conn)
            except OSError:
                pass
            finally:
                self.connected = False
                self._log("Client disconnected.")
                try:
                    conn.close()
                except OSError:
                    pass

        if self._sock:
            try:
                self._sock.close()
            except OSError:
                pass

    def _comm_loop(self, conn: socket.socket):
        conn.settimeout(0.5)
        buf = ""
        while self.running:
            try:
                data = conn.recv(4096)
                if not data:
                    break
                buf += data.decode("ascii", errors="ignore")
            except socket.timeout:
                continue

            while "\n" in buf:
                line, buf = buf.split("\n", 1)
                line = line.strip("\r\n ")
                if not line:
                    continue
                self._log(f"RX: {line}")
                response = self._handle_command(line)
                if response is not None:
                    self._log(f"TX: {response}")
                    conn.sendall((response + "\n").encode("ascii"))

    # ── SCPI command handling ────────────────────────────────────────────────

    def _handle_command(self, line: str) -> str | None:
        upper = line.upper()

        if upper == "*IDN?":
            return self.state.idn

        if upper.startswith("CONF:"):
            self.state.configure(line[len("CONF:"):])
            return None

        if upper.startswith("SAMP:COUN"):
            parts = line.split(None, 1)
            if len(parts) == 2:
                try:
                    self.state.set_sample_count(int(float(parts[1].strip())))
                except ValueError:
                    pass
            return None

        if upper.startswith("CALC:AVER:STAT"):
            parts = line.split(None, 1)
            if len(parts) == 2:
                self.state.set_averaging(parts[1].strip().upper() in ("ON", "1"))
            return None

        if upper in ("INIT", "INIT:IMM"):
            self.state.trigger()
            return None

        if upper == "CALC:AVER:COUN?":
            return str(self.state.current_aver_count())

        if upper == "CALC:AVER:ALL?":
            return f"{self.state.averaged_reading():.6f}"

        if upper.endswith("?"):
            return "0"

        return None

    def _log(self, msg: str):
        self.log_queue.put(msg)
