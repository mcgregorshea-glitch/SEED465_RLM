import threading
import serial
import queue
import time
from marlin_logic import MarlinState, GCodeParser

class SerialServer:
    """
    Handles serial communication in a background thread.
    Parses incoming G-Code and sends responses.
    """
    def __init__(self, port: str, baudrate: int = 115200):
        self.port = port
        self.baudrate = baudrate
        self.running = False
        self.thread = None
        
        self.state = MarlinState()
        self.parser = GCodeParser(self.state)
        
        # Queues for sending data to the GUI
        self.update_queue = queue.Queue()
        self.log_queue = queue.Queue()
        
    def start(self):
        """Starts the serial listener thread."""
        self.running = True
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()

    def stop(self):
        """Stops the serial listener thread."""
        self.running = False
        if self.thread:
            self.thread.join(timeout=1.0)

    def _run(self):
        """Main loop for the serial thread."""
        try:
            ser_inst = None
            try:
                ser_inst = serial.Serial(self.port, self.baudrate, timeout=1)
                print(f"SerialServer listening on {self.port} at {self.baudrate} baud.")
            except serial.SerialException:
                print(f"WARNING: Could not open {self.port}. Running in MOCK mode.")
            
            if ser_inst:
                ser_inst.write(b"start\necho: Marlin 1.1.9\nok\n")
            
            buffer = ""
            last_sim_time = time.perf_counter()
            
            while self.running:
                # 1. Physical Simulation Step
                now = time.perf_counter()
                dt = min(now - last_sim_time, 0.1)
                last_sim_time = now
                
                if dt > 0:
                    self.state.step_motion(dt, speed_factor=getattr(self, 'speed_factor', 1.0))
                
                # 2. Serial Communication
                if ser_inst and ser_inst.in_waiting > 0:
                    try:
                        data = ser_inst.read(ser_inst.in_waiting).decode('utf-8', errors='ignore')
                        buffer += data
                    except Exception as e:
                        print(f"Serial Read Error: {e}")
                        continue
                    
                    if "\n" in buffer or "\r" in buffer:
                        # Split by any line ending, keeping partial last line
                        lines = buffer.splitlines(keepends=True)
                        if not buffer.endswith(("\n", "\r")):
                            buffer = lines.pop()
                        else:
                            buffer = ""
                        
                        for raw_line in lines:
                            line = raw_line.strip()
                            if not line: continue
                            
                            print(f"[SERIAL RX] {line}")
                            self.log_queue.put(("RX", line))
                            response = self.parser.parse_line(line)

                            # Blocking Handshake Logic
                            cmd_parts = line.split()
                            if cmd_parts:
                                cmd_type = cmd_parts[0].upper()
                                if cmd_type in ['G0', 'G1', 'G28'] and response == "ok":
                                    while not self.state.at_target and self.running:
                                        now = time.perf_counter()
                                        dt = min(now - last_sim_time, 0.05)
                                        last_sim_time = now
                                        if dt > 0:
                                            self.state.step_motion(dt, speed_factor=getattr(self, 'speed_factor', 1.0))
                                        time.sleep(0.001)

                            print(f"[SERIAL TX] {response}")
                            ser_inst.write((response + "\n").encode('utf-8'))
                            self.log_queue.put(("TX", response))
                
                time.sleep(0.005)
                
            if ser_inst:
                ser_inst.close()

        except Exception as e:
            print(f"Serial Error: {e}")
            self.running = False
