from typing import Dict
import math

class MarlinState:
    """
    Maintains the internal state of a simulated Marlin-based 3D printer.
    Handles coordinate tracking (X, Y, Z, E) and positioning modes.
    """
    def __init__(self):
        # Current position
        self.x: float = 0.0
        self.y: float = 0.0
        self.z: float = 0.0
        self.e: float = 0.0
        
        # Target position (where we are moving to)
        self.target_x: float = 0.0
        self.target_y: float = 0.0
        self.target_z: float = 0.0
        self.target_e: float = 0.0
        
        self.feedrate: float = 1500.0  # mm/min (standard Marlin default)
        
        self.absolute_positioning: bool = True  # G90 (True) / G91 (False)
        self.absolute_extrusion: bool = True    # M82 (True) / M83 (False)
        
        # Bed bounds for an Ender 3
        self.min_x, self.max_x = 0.0, 220.0
        self.min_y, self.max_y = 0.0, 220.0
        self.min_z, self.max_z = 0.0, 250.0

    def get_state(self) -> Dict[str, float]:
        """Returns the current coordinates."""
        return {
            'X': self.x,
            'Y': self.y,
            'Z': self.z,
            'E': self.e
        }

    def update_position(self, x=None, y=None, z=None, e=None):
        """Immediately sets current position (with clamping, respecting positioning mode)."""
        if self.absolute_positioning:
            if x is not None: self.x = self.target_x = self._clamp(x, self.min_x, self.max_x)
            if y is not None: self.y = self.target_y = self._clamp(y, self.min_y, self.max_y)
            if z is not None: self.z = self.target_z = self._clamp(z, self.min_z, self.max_z)
        else:
            if x is not None: self.x = self.target_x = self._clamp(self.x + x, self.min_x, self.max_x)
            if y is not None: self.y = self.target_y = self._clamp(self.y + y, self.min_y, self.max_y)
            if z is not None: self.z = self.target_z = self._clamp(self.z + z, self.min_z, self.max_z)
        if e is not None:
            if self.absolute_extrusion:
                self.e = self.target_e = float(e)
            else:
                self.e = self.target_e = self.e + float(e)

    def set_target(self, x=None, y=None, z=None, e=None, f=None):
        """Sets the next destination for the machine."""
        if f is not None:
            self.feedrate = f

        if self.absolute_positioning:
            if x is not None: self.target_x = self._clamp(x, self.min_x, self.max_x)
            if y is not None: self.target_y = self._clamp(y, self.min_y, self.max_y)
            if z is not None: self.target_z = self._clamp(z, self.min_z, self.max_z)
        else:
            if x is not None: self.target_x = self._clamp(self.target_x + x, self.min_x, self.max_x)
            if y is not None: self.target_y = self._clamp(self.target_y + y, self.min_y, self.max_y)
            if z is not None: self.target_z = self._clamp(self.target_z + z, self.min_z, self.max_z)

        if e is not None:
            if self.absolute_extrusion:
                self.target_e = e
            else:
                self.target_e += e

    def step_motion(self, dt: float, speed_factor: float = 1.0):
        """
        Moves current position toward target based on feedrate and dt.
        dt is in seconds. feedrate is in mm/min.
        """
        # mm per second
        v = (self.feedrate / 60.0) * speed_factor
        
        dx = self.target_x - self.x
        dy = self.target_y - self.y
        dz = self.target_z - self.z
        de = self.target_e - self.e
        
        dist = math.sqrt(dx**2 + dy**2 + dz**2)
        
        if dist < (v * dt) or dist == 0:
            # Arrived or very close
            self.x, self.y, self.z = self.target_x, self.target_y, self.target_z
            self.e = self.target_e # E is usually handled separately but simplified here
        else:
            # Move along vector
            ratio = (v * dt) / dist
            self.x += dx * ratio
            self.y += dy * ratio
            self.z += dz * ratio
            # Simple linear E interpolation
            self.e += de * ratio

    def home(self):
        """Resets target and current to 0, 0, 0."""
        self.x = self.y = self.z = 0.0
        self.target_x = self.target_y = self.target_z = 0.0

    @property
    def at_target(self) -> bool:
        """Returns True if the current position matches the target position."""
        return (
            abs(self.x - self.target_x) < 0.001 and
            abs(self.y - self.target_y) < 0.001 and
            abs(self.z - self.target_z) < 0.001 and
            abs(self.e - self.target_e) < 0.001
        )

    def _clamp(self, val: float, min_val: float, max_val: float) -> float:
        return max(min_val, min(val, max_val))

    def format_m114(self) -> str:
        """Formats the response for an M114 command."""
        return f"X:{self.x:.2f} Y:{self.y:.2f} Z:{self.z:.2f} E:{self.e:.2f} Count X:0 Y:0 Z:0"

class GCodeParser:
    """
    Parses G-Code strings and updates a MarlinState instance.
    """
    def __init__(self, state: MarlinState):
        self.state = state

    def parse_line(self, line: str) -> str:
        """
        Parses a single line of G-Code.
        Returns the response string (e.g., 'ok' or position info).
        """
        # Remove comments and whitespace
        clean_line = line.split(';')[0].strip()
        if not clean_line:
            return "ok"

        parts = clean_line.split()
        command = parts[0].upper()
        params = self._parse_params(parts[1:])

        response = "ok"

        if command in ['G0', 'G1']:
            if params.get('F') is not None:
                self.state.feedrate = params['F']
            self.state.update_position(
                x=params.get('X'),
                y=params.get('Y'),
                z=params.get('Z'),
                e=params.get('E'),
            )
        elif command == 'G28':
            self.state.home()
        elif command == 'G90':
            self.state.absolute_positioning = True
        elif command == 'G91':
            self.state.absolute_positioning = False
        elif command == 'M82':
            self.state.absolute_extrusion = True
        elif command == 'M83':
            self.state.absolute_extrusion = False
        elif command == 'M114':
            response = self.state.format_m114() + "\nok"
        elif command == 'M117':
            # LCD Message - just acknowledge
            pass
        elif command == 'M112':
            # Emergency Stop - could reset state or set a flag
            self.state.home()
            response = "Error:Emergency Stop\nok"

        return response

    def _parse_params(self, param_list: list) -> Dict[str, float]:
        """Extracts X, Y, Z, E values from G-Code parameters."""
        params = {}
        for p in param_list:
            if len(p) < 2: continue
            key = p[0].upper()
            try:
                val = float(p[1:])
                params[key] = val
            except ValueError:
                continue
        return params
