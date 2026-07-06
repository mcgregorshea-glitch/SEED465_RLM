import time
import threading
import serial
import serial.tools.list_ports
from typing import Dict, List, Optional, Any, Callable, Tuple
import re

class SerialEngine:
    """
    Manages the serial connection to the printer and handles the background 
    threads for connecting and sending G-code.
    """
    def __init__(self, message_queue, stop_event, pause_event) -> None:
        self.message_queue = message_queue
        self.stop_event = stop_event
        self.pause_event = pause_event
        self.connection: Optional[serial.Serial] = None
        self.port: Optional[str] = None
        self.baudrate: int = 115200
        self.lock = threading.Lock()
        
        # Callbacks for GUI interactions
        self.callbacks = {
            'take_measurement': None,
            'homing_verification': None,
            'handle_homing_failure': None,
            'get_hub_panel': None,
            'apply_e_conversion': lambda x: x,
            'parse_coords': None,
            'get_speed_cap': lambda: 99999,
            'set_speed_cap': lambda x: None
        }

    def set_callback(self, name: str, func: Callable) -> None:
        self.callbacks[name] = func

    def queue_message(self, text: str, level: str = "INFO") -> None:
        self.message_queue.put(("LOG", (text, level)))

    def is_connected(self) -> bool:
        with self.lock:
            return self.connection is not None and self.connection.is_open

    def write_command(self, cmd: str) -> None:
        """Sends a raw command string to the printer thread-safely."""
        with self.lock:
            if self.connection and self.connection.is_open:
                try:
                    self.queue_message(f"<- {cmd}")
                    self.connection.write(cmd.encode('utf-8') + b'\n')
                    self.connection.flush()
                except Exception as e:
                    self.queue_message(f"Serial Write Error: {e}", "ERROR")

    def connect(self, selected_port: str, baudrate: int, cancel_event: threading.Event) -> None:
        """Starts the connection thread."""
        threading.Thread(
            target=self._connect_thread, 
            args=(selected_port, baudrate, cancel_event), 
            daemon=True
        ).start()

    def _connect_thread(self, selected_port: str, baudrate: int, cancel_event: threading.Event) -> None:
        """Worker thread for establishing serial connection."""
        available_ports = [p.device for p in serial.tools.list_ports.comports()]
        ports_to_try = []
        if selected_port == "Auto-detect":
            self.queue_message("Auto-detecting printer port...")
            ports_to_try = available_ports
            if not ports_to_try:
                self.queue_message("No serial ports found.", "ERROR")
                self.message_queue.put(("CONNECT_FAIL", "No serial ports found."))
                return
        else:
            ports_to_try = [selected_port]

        serial_conn, found_port, connection_cancelled = None, None, False
        try:
            for port in ports_to_try:
                if cancel_event.is_set():
                    connection_cancelled = True
                    break
                
                self.queue_message(f"Trying port: {port}...")
                temp_ser = None
                try:
                    temp_ser = serial.Serial(port=port, baudrate=baudrate, timeout=5, write_timeout=10, dsrdtr=False, rtscts=False)
                    try:
                        temp_ser.setDTR(False)
                        time.sleep(0.02)
                        temp_ser.setRTS(False)
                    except:
                        pass
                        
                    self.queue_message(f"Port {port} opened. Waiting...")
                    time.sleep(5)
                    temp_ser.reset_input_buffer()
                    time.sleep(1.5)
                    
                    # Verify responsiveness
                    responsive = False
                    keywords = ['ok', 't:', 'temp', 'echo:', 'marlin', 'start', 'wait']
                    for attempt in range(3):
                        if cancel_event.is_set(): raise InterruptedError("Cancelled")
                        self.queue_message(f"Sending M105 to {port} ({attempt + 1})...")
                        temp_ser.write(b'M105\n')
                        temp_ser.flush()
                        
                        start_time = time.time()
                        while time.time() - start_time < 5.0:
                            if cancel_event.is_set(): raise InterruptedError("Cancelled")
                            if temp_ser.in_waiting > 0:
                                line = temp_ser.readline().decode('utf-8', errors='ignore').strip()
                                if line: self.queue_message(f"Resp on {port}: {line}")
                                if any(k in line.lower() for k in keywords):
                                    responsive = True
                                    break
                            time.sleep(0.1)
                        if responsive: break
                        time.sleep(1.0)
                        
                    if responsive:
                        serial_conn, found_port = temp_ser, port
                        break
                    else:
                        temp_ser.close()
                except InterruptedError:
                    connection_cancelled = True
                    break
                except Exception as e:
                    self.queue_message(f"Error on {port}: {e}", "WARN")
                    if temp_ser: temp_ser.close()

            if connection_cancelled:
                self.message_queue.put(("CONNECT_CANCELLED", None))
            elif serial_conn and found_port:
                serial_conn.write(b'M302 P1 S0\nM84 S900\nM120\n')
                serial_conn.flush()
                time.sleep(1) 
                
                initial_position = self._poll_position(serial_conn)
                self.message_queue.put(("CONNECTED", (serial_conn, found_port, baudrate, initial_position)))
                with self.lock:
                    self.connection = serial_conn
                    self.port = found_port
                    self.baudrate = baudrate
            else:
                self.message_queue.put(("CONNECT_FAIL", "No responsive printer found."))
        finally:
            self.message_queue.put(("CONNECT_ATTEMPT_FINISHED", None))

    def _poll_position(self, conn: serial.Serial) -> Optional[Dict[str, float]]:
        conn.write(b'M114\n')
        conn.flush()
        start = time.time()
        buf = ""
        while time.time() - start < 5.0:
            if conn.in_waiting > 0:
                buf += conn.read(conn.in_waiting).decode('utf-8', errors='ignore')
                if 'ok' in buf.lower(): break
            time.sleep(0.05)
        
        return self._parse_m114(buf)

    @staticmethod
    def _parse_m114(buf: str) -> Optional[Dict[str, float]]:
        """Parses an M114 reply into {x, y, z, rot}. The first X:/Y:/Z:/E: token
        is the logical position (a trailing 'Count X:.. Y:.. Z:..' block reports
        stepper counts and is ignored — re.search takes the first match)."""
        matches = {
            'x': re.search(r'X:\s*([-+]?\d*\.?\d+)', buf),
            'y': re.search(r'Y:\s*([-+]?\d*\.?\d+)', buf),
            'z': re.search(r'Z:\s*([-+]?\d*\.?\d+)', buf),
            'rot': re.search(r'E:\s*([-+]?\d*\.?\d+)', buf)  # 'rot' = unified 4th axis key
        }
        if matches['x'] and matches['y'] and matches['z']:
            pos = {k: float(m.group(1)) for k, m in matches.items() if m}
            if 'rot' not in pos: pos['rot'] = 0.0
            return pos
        return None

    def _read_actual_position(self, fallback: Dict[str, float]) -> Dict[str, float]:
        """Queries the printer's ACTUAL position (M114) mid-scan and returns it.

        Measurements must be logged against where the head physically is, not
        where it was commanded to go: if the firmware clamps an out-of-range
        target (or the axis skips steps), the two differ and logging the target
        silently records a wrong coordinate. Falls back to the commanded target
        if the reply can't be read/parsed so logging never breaks the scan.
        """
        try:
            with self.lock:
                if not self.connection: return fallback
                self.connection.write(b'M114\n')
            start = time.time()
            buf = ""
            while time.time() - start < 5.0:
                if self.stop_event.is_set(): break
                with self.lock:
                    if self.connection and self.connection.in_waiting > 0:
                        buf += self.connection.read(self.connection.in_waiting).decode('utf-8', errors='ignore')
                if 'ok' in buf.lower(): break
                time.sleep(0.02)
            parsed = self._parse_m114(buf)
            if parsed is not None:
                # Preserve the commanded rot if the reply carried no E word.
                if 'rot' not in parsed or parsed.get('rot') is None:
                    parsed['rot'] = fallback.get('rot', 0.0)
                return parsed
        except Exception as e:
            self.queue_message(f"M114 position read failed, logging commanded target: {e}", "WARN")
        return fallback

    def _wait_for_ok(self, timeout_s=30):
        start_t = time.time()
        buf = ""
        while time.time() - start_t < timeout_s:
            if self.stop_event.is_set(): raise InterruptedError("Stop")
            with self.lock:
                if self.connection and self.connection.in_waiting > 0:
                    resp = self.connection.read(self.connection.in_waiting).decode('utf-8', errors='ignore')
                    buf += resp
                    for line in resp.splitlines():
                        line = line.strip()
                        if line: self.queue_message(f"-> {line}")
                    if 'ok' in buf.lower(): return True
            time.sleep(0.05)
        return False

    def _wait_for_n_oks(self, count: int, timeout_s: int = 30) -> bool:
        """Wait until `count` total 'ok' acknowledgements have been read.

        Marlin batches replies when commands are written back-to-back, so a
        single serial read can contain several 'ok's at once. The old pattern
        of chaining two `_wait_for_ok()` calls starved the second waiter: the
        first call consumed every 'ok' in the buffer and returned, leaving the
        second to burn its full timeout (the ~30s recovery stall). Counting
        across one accumulating buffer avoids that.
        """
        start_t = time.time()
        buf = ""
        while time.time() - start_t < timeout_s:
            if self.stop_event.is_set(): raise InterruptedError("Stop")
            with self.lock:
                if self.connection and self.connection.in_waiting > 0:
                    resp = self.connection.read(self.connection.in_waiting).decode('utf-8', errors='ignore')
                    buf += resp
                    for line in resp.splitlines():
                        line = line.strip()
                        if line: self.queue_message(f"-> {line}")
            seen = sum(1 for ln in buf.lower().splitlines() if ln.strip().startswith('ok'))
            if seen >= count: return True
            time.sleep(0.05)
        return False

    def _apply_speed_cap(self, line: str) -> str:
        """Clamp the feedrate (F word) of a move to the active speed cap.

        The drift-recovery logic lowers the cap after repeated skips, but the
        value was only stored and logged — never sent. Without this clamp the
        scan "resumes at reduced speed" in name only. The default sentinel
        (99999) means no cap. A move with no explicit F gets the cap appended
        so it can't inherit a faster prior feedrate.
        """
        cap = self.callbacks['get_speed_cap']()
        if cap is None or cap >= 99999:
            return line
        m = re.search(r'[Ff](\d+(?:\.\d+)?)', line)
        if m:
            if float(m.group(1)) > cap:
                return line[:m.start()] + f"F{cap:.0f}" + line[m.end():]
            return line
        return f"{line} F{cap:.0f}"

    def send_gcode(self, steps: List[Any], initial_pos: Dict[str, float], bounds: Dict[str, float], auto_measure: bool) -> None:
        """Starts the sender thread."""
        threading.Thread(
            target=self._sender_thread,
            args=(steps, initial_pos, bounds, auto_measure),
            daemon=True
        ).start()

    def _sender_thread(self, steps_to_run: List[Any], initial_pos: Dict[str, float], bounds: Dict[str, float], auto_measure: bool) -> None:
        """The main background worker thread for sending G-code."""
        self.message_queue.put(("SET_STATUS", "on"))
        # Start every scan at full speed; the cap is persistent on the panel and
        # would otherwise carry a lowered value over from a previous drift run.
        self.callbacks['set_speed_cap'](99999)
        total_steps = len(steps_to_run)
        sent_step_count = 0
        success = True
        moves_sent = 0
        last_pos = initial_pos.copy()

        cmd_index = 0
        layer_start_index = 0
        layer_start_moves = 0
        layer_retry_count = 0
        first_layer_change = True  # skip drift verification on the first layer
                                   # boundary only (printer was just homed at
                                   # scan start); every later boundary is checked
        just_rehomed = False  # set after a blind recovery G28 so the rewound
                              # boundary move resumes instead of re-verifying
        _pre_measure_executed = set()  # tracks cmd_indices whose look-ahead has already run

        while cmd_index < len(steps_to_run):
            self.pause_event.wait()
            if self.stop_event.is_set():
                success = False
                break

            step = steps_to_run[cmd_index]
            
            # Advanced Sequence Steps (Hub Commands / Waits)
            if isinstance(step, dict):
                stype = step.get('type')
                if stype == 'hub_cmd':
                    param, value = step.get('param'), step.get('value')
                    self.queue_message(f"Sequence: Sending Hub command {param}={value}")
                    hub_panel = self.callbacks['get_hub_panel']() if self.callbacks['get_hub_panel'] else None
                    if hub_panel and getattr(hub_panel, 'rlm', None):
                        try: hub_panel.rlm.set_param(param, value)
                        except Exception as e: self.queue_message(f"Hub Error: {e}", "ERROR")
                    else: self.queue_message("Hub not connected, skipping.", "WARN")
                elif stype == 'wait':
                    secs = step.get('seconds', 0)
                    self.queue_message(f"Sequence: Waiting {secs}s...")
                    st = time.time()
                    while time.time() - st < secs:
                        if self.stop_event.is_set(): break
                        time.sleep(0.1)
                # Skip pre-measure steps already executed by the look-ahead.
                # Not counted toward sent_step_count so the progress bar advances
                # smoothly with G1 moves rather than jumping at pre-measure steps.
                elif stype in ('pre_measure_hub_cmd', 'pre_measure_wait'):
                    cmd_index += 1; continue
                # --- END NEW ---
                elif stype == 'move':
                    gcode_line = step.get('gcode', "")
                else:
                    self.queue_message(f"Unknown step type: {stype}", "WARN")
                    cmd_index += 1; continue
                
                if stype != 'move':
                    cmd_index += 1; sent_step_count += 1; continue
            else:
                gcode_line = str(step).strip()

            if not gcode_line or gcode_line.startswith(';'):
                cmd_index += 1; continue

            is_move = gcode_line.upper().startswith("G0") or gcode_line.upper().startswith("G1")
            
            # Progress Update
            sent_step_count += 1
            perc = (sent_step_count / total_steps) * 100 if total_steps > 0 else 0
            self.message_queue.put(("PROGRESS", (sent_step_count, total_steps, perc)))

            target_pos = None
            if is_move:
                parsed = self.callbacks['parse_coords'](gcode_line) if self.callbacks['parse_coords'] else {}
                if parsed:
                    target_pos = last_pos.copy()
                    target_pos.update(parsed)
                    
                    # Layer change detection
                    if target_pos['z'] != last_pos['z'] and last_pos['z'] is not None:
                        if moves_sent > 0 and first_layer_change:
                            # First completed layer: printer was homed at scan
                            # start, so no drift check is needed. Consume the
                            # one-shot so every later boundary IS verified.
                            # (The old `last_pos['z'] == z_min` proxy re-skipped
                            # every departure from the bottom layer, silently
                            # dropping verification on each rotation's first
                            # layer in multi-rotation scans.)
                            first_layer_change = False
                        elif moves_sent > 0:
                            if cmd_index != layer_start_index:
                                layer_start_index = cmd_index
                                layer_start_moves = moves_sent

                            if just_rehomed:
                                # The recovery G28 below already synced the
                                # coordinate system, so trust this boundary move
                                # and resume the scan (at the reduced cap) rather
                                # than re-verifying — re-verifying here is what
                                # produced the endless verify/home/verify loop.
                                just_rehomed = False
                            else:
                                self.queue_message(f"Layer change ({last_pos['z']:.2f} -> {target_pos['z']:.2f}). Verifying...")
                                try:
                                    if self.callbacks['homing_verification']:
                                        self.callbacks['homing_verification'](auto_restart=True)
                                    layer_retry_count = 0
                                except ValueError as e:
                                    layer_retry_count += 1
                                    diag = str(e)
                                    if layer_retry_count >= 3:
                                        current_cap = self.callbacks['get_speed_cap']()
                                        new_cap = 1000 if current_cap > 1000 else current_cap / 2
                                        self.callbacks['set_speed_cap'](new_cap)
                                        self.queue_message(f"STABILITY ALERT: 3 drifts. Speed cap: {new_cap:.0f}", "WARN")
                                        layer_retry_count = 0
                                        if new_cap <= 100:
                                            # Even at the floor speed the machine
                                            # keeps skipping — stop looping and
                                            # hand off to the failure handler.
                                            if self.callbacks['handle_homing_failure']:
                                                self.callbacks['handle_homing_failure'](diag)
                                            raise InterruptedError(diag)

                                    self.queue_message(f"DRIFT (Attempt {layer_retry_count}/3): {diag}", "WARN")
                                    try:
                                        with self.lock:
                                            self.connection.write(b'G28 X Y\nM400\n')
                                        self._wait_for_n_oks(2, 90)
                                        just_rehomed = True
                                        cmd_index = layer_start_index
                                        moves_sent = layer_start_moves
                                        self.message_queue.put(("PATH_PROGRESS_UPDATE", moves_sent))
                                        continue
                                    except Exception as recovery_err:
                                        if self.callbacks['handle_homing_failure']:
                                            self.callbacks['handle_homing_failure'](diag)
                                        raise InterruptedError(diag)

            elif "G28" in gcode_line.upper():
                target_pos = last_pos.copy()
                if 'X' not in gcode_line.upper() and 'Y' not in gcode_line.upper() and 'Z' not in gcode_line.upper():
                    target_pos.update({'x': bounds['x_min'], 'y': bounds['y_min'], 'z': bounds['z_min']})
                else:
                    if 'X' in gcode_line.upper(): target_pos['x'] = bounds['x_min']
                    if 'Y' in gcode_line.upper(): target_pos['y'] = bounds['y_min']
                    if 'Z' in gcode_line.upper(): target_pos['z'] = bounds['z_min']
                if 'E' in gcode_line.upper(): target_pos['rot'] = 0.0

            # Homing Verification Intercept
            if "G28" in gcode_line.upper() and 'X' in gcode_line.upper() and 'Y' in gcode_line.upper() and 'Z' not in gcode_line.upper():
                if sent_step_count > 1:
                    self.queue_message("Explicit G28 X Y intercepted. Verifying...")
                    try:
                        if self.callbacks['homing_verification']:
                            self.callbacks['homing_verification'](auto_restart=True)
                        last_pos.update(target_pos)
                        cmd_index += 1; continue
                    except ValueError as e:
                        # Recovery logic same as above
                        diag = str(e)
                        self.queue_message(f"DRIFT: {diag}", "WARN")
                        try:
                            with self.lock: self.connection.write(b'G28 X Y\nM400\n')
                            self._wait_for_n_oks(2, 90)
                            just_rehomed = True
                            cmd_index = layer_start_index
                            moves_sent = layer_start_moves
                            self.message_queue.put(("PATH_PROGRESS_UPDATE", moves_sent))
                            continue
                        except:
                            if self.callbacks['handle_homing_failure']: self.callbacks['handle_homing_failure'](diag)
                            raise InterruptedError(diag)

            # Send and Wait
            try:
                line_to_send = self.callbacks['apply_e_conversion'](gcode_line)
                if is_move:
                    line_to_send = self._apply_speed_cap(line_to_send)
                with self.lock:
                    if not self.connection: break
                    self.queue_message(f"<- {line_to_send}")
                    self.connection.write(line_to_send.encode('utf-8') + b'\n')
                
                if self._wait_for_ok():
                    if target_pos:
                        last_pos = target_pos
                        self.message_queue.put(("POSITION_UPDATE", last_pos.copy()))
                    if is_move or "G28" in gcode_line.upper():
                        moves_sent += 1
                        self.message_queue.put(("PATH_PROGRESS_UPDATE", moves_sent))
                    
                    if is_move and auto_measure:
                        with self.lock: self.connection.write(b'M400\n')
                        if self._wait_for_ok():
                            # --- Execute pre-measurement steps (look-ahead) ---
                            # Guard: only run once per cmd_index; FORCE_PAUSE can re-enter
                            # with the same cmd_index, and we must not double-fire hub cmds.
                            if cmd_index not in _pre_measure_executed:
                                _pre_measure_executed.add(cmd_index)
                                peek = cmd_index + 1
                                while peek < len(steps_to_run):
                                    s = steps_to_run[peek]
                                    if not isinstance(s, dict):
                                        break
                                    stype_peek = s.get('type')
                                    if stype_peek == 'pre_measure_hub_cmd':
                                        hub_panel = self.callbacks['get_hub_panel']() if self.callbacks.get('get_hub_panel') else None
                                        if hub_panel and getattr(hub_panel, 'rlm', None):
                                            try:
                                                self.queue_message(f"Pre-measure: {s.get('param')}={s.get('value')}")
                                                hub_panel.rlm.set_param(s.get('param'), s.get('value'))
                                            except Exception as e:
                                                self.queue_message(f"Pre-measure Hub Error: {e}", "ERROR")
                                        else:
                                            self.queue_message("Hub not connected, skipping pre-measure cmd.", "WARN")
                                        peek += 1
                                    elif stype_peek == 'pre_measure_wait':
                                        secs = s.get('seconds', 0)
                                        st = time.time()
                                        while time.time() - st < secs:
                                            if self.stop_event.is_set():
                                                break
                                            time.sleep(0.1)
                                        peek += 1
                                        if self.stop_event.is_set():
                                            break  # abort look-ahead; stop_event checked below
                                    else:
                                        break
                            # --- End pre-measurement look-ahead ---
                            if self.callbacks['take_measurement']:
                                # Log against the printer's ACTUAL position, not
                                # the commanded target — a firmware travel clamp
                                # or skipped steps would otherwise be recorded as
                                # the intended coordinate. The head has settled
                                # here (M400 above), so one read holds for all
                                # measurement retries at this point.
                                measured_pos = self._read_actual_position(last_pos)
                                # Gated Measurement: Retry until we get valid data
                                vals = []
                                retries = 0
                                max_retries = 60 # 2 minute total tolerance
                                while retries < max_retries:
                                    vals = self.callbacks['take_measurement']()
                                    if vals: # Non-empty list
                                        self.message_queue.put(("MEASUREMENT_RESULT", (vals, measured_pos.copy())))
                                        break
                                    
                                    if self.stop_event.is_set(): break
                                    retries += 1
                                    if retries % 5 == 0:
                                        self.queue_message(f"Waiting for telemetry data... (Attempt {retries}/{max_retries})", "WARN")
                                    time.sleep(2.0)
                                
                                if not vals and not self.stop_event.is_set():
                                    self.queue_message("Gated Measurement: Failed to capture data after 60 attempts. Pausing scan.", "ERROR")
                                    self.pause_event.clear()
                                    self.message_queue.put(("FORCE_PAUSE", True))
                                    continue # Re-run this cmd_index when resumed
                    cmd_index += 1
                else:
                    self.queue_message(f"No 'ok' for '{gcode_line}'", "WARN")
                    success = False; break
            except Exception as e:
                self.queue_message(f"Error: {e}", "ERROR")
                success = False; break

        self.message_queue.put(("PROGRESS_RESET", None))
        self.message_queue.put(("FILE_SEND_FINISHED", None))
